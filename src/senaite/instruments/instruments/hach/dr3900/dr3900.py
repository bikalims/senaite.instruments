# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS.
#
# SENAITE.CORE is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, version 2.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 51
# Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#
# Copyright 2018-2019 by it's authors.
# Some rights reserved, see README and LICENSE.
import re
import csv
import json
import traceback
from mimetypes import guess_type
from os.path import abspath
from os.path import splitext
from DateTime import DateTime
from bika.lims.browser import BrowserView

from senaite.core.exportimport.instruments import (
    IInstrumentAutoImportInterface, IInstrumentImportInterface
)
from senaite.core.exportimport.instruments import IInstrumentExportInterface
from senaite.core.exportimport.instruments.resultsimport import (
    AnalysisResultsImporter)
from senaite.core.exportimport.instruments.resultsimport import (
    InstrumentResultsFileParser)
from senaite.instruments.instrument import xls_to_csv
from senaite.instruments.instrument import xlsx_to_csv

from bika.lims import api
from bika.lims import bikaMessageFactory as _
from bika.lims.catalog import CATALOG_ANALYSIS_REQUEST_LISTING
from senaite.core.catalog import ANALYSIS_CATALOG, SENAITE_CATALOG
from senaite.instruments.instrument import FileStub
from senaite.instruments.instrument import SheetNotFound
from zope.interface import implements
from zope.publisher.browser import FileUpload

field_interim_map = {"Dilution": "Factor","Result": "Reading"}


class SampleNotFound(Exception):
    pass


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class DR3900Parser(InstrumentResultsFileParser):
    ar = None

    def __init__(self, infile, worksheet=None, encoding=None, delimiter=None):
        self.delimiter = delimiter if delimiter else ","
        self.encoding = encoding
        self.ar = None
        self.analyses = None
        self.worksheet = worksheet if worksheet else 0
        self.infile = infile
        self.csv_data = None
        self.sample_id = None
        self.processed_samples = []
        mimetype, encoding = guess_type(self.infile.filename)
        InstrumentResultsFileParser.__init__(self, infile, mimetype)


    def parse(self):
        order = []
        ext = splitext(self.infile.filename.lower())[-1]
        if ext == ".xlsx": #fix in flameatomic also
            order = (xlsx_to_csv, xls_to_csv)
        elif ext == ".xls":
            order = (xls_to_csv, xlsx_to_csv)
        elif ext == ".csv" or ext == ".prn":
            self.csv_data = self.infile
        if order:
            for importer in order:
                try:
                    self.csv_data = importer(
                        infile=self.infile,
                        worksheet=self.worksheet,
                        delimiter=self.delimiter,
                    )
                    break
                except SheetNotFound:
                    self.err("Sheet not found in workbook: %s" % self.worksheet)
                    return -1
                except Exception as e:
                    pass
            else:
                self.warn("Can't parse input file as XLS, XLSX, or CSV.")
                return -1

        stub = FileStub(file=self.csv_data, name=str(self.infile.filename))
        self.csv_data = FileUpload(stub)
        data = self.csv_data.read()

        decoded_data = self.try_utf8(data)
        if decoded_data:
            if ext == ".xlsx":
                lines_with_parentheses = decoded_data.split("\n")
            else:
                lines_with_parentheses = decoded_data.split("\r\n")
        else:
            decoded_data = self.try_utf16(data)
            if decoded_data:
                if "\r\n" in decoded_data:
                    lines_with_parentheses = data.decode('utf-16').split("\r\n")
                elif "\n" in decoded_data:
                    lines_with_parentheses = data.decode('utf-16').split("\n")
                else:
                    lines_with_parentheses = re.sub(r'[^\x00-\x7f]',r'', data).split("\r\n") #if bad conversion
            else:
                if "\r\n" in data:
                    lines_with_parentheses = re.sub(r'[^\x00-\x7f]',r'', data).split("\r\n")
                else:
                    lines_with_parentheses = re.sub(r'[^\x00-\x7f]',r'', data).split("\n")
        lines = [i.replace('"','') for i in lines_with_parentheses]
        
        ascii_lines = self.extract_relevant_data(lines)
        reader = csv.DictReader(ascii_lines)

        headers_parsed = self.parse_headerlines(reader)

        if headers_parsed:
            for row in reader:
                self.parse_row(row,reader.line_num)
        return 0


    def parse_row(self, row, row_nr):
        parsed_strings = {}

        parsed_strings = self.interim_map_sorter(row)
        parsed = self.data_cleaning(parsed_strings)
        sample_ID = row.get("Sample ID:")
        regex = re.compile('[^a-zA-Z]')
        sample_service = regex.sub('',row.get("Parameter:"))

        if not sample_service or not sample_ID or not row.get("Result").strip(" "):
            self.warn("Data not entered correctly for '{}' with sample ID '{}' and result of '{}'".format(sample_service,sample_ID,row.get("Result")))
            return 0

        if {sample_ID:sample_service} in self.processed_samples:
            msg = ("Multiple results for Sample '{}' with sample service '{}' found. Not imported".format(sample_ID,sample_service))
            raise MultipleAnalysesFound(msg)

        try:
            if self.is_sample(sample_ID):
                ar = self.get_ar(sample_ID)
                analysis = self.get_analysis(ar,sample_service)
                keyword = analysis.getKeyword
            elif self.is_analysis_group_id(sample_ID):
                analysis = self.get_duplicate_or_qc(sample_ID,sample_service)
                keyword = analysis.getKeyword
            else:
                sample_reference = self.get_reference_sample(sample_ID, sample_service)
                analysis = self.get_reference_sample_analysis(sample_reference, sample_service)
                keyword = analysis.getKeyword()
        except Exception as e:
            self.warn(msg="Error getting analysis for '${s}/${kw}': ${e}",
                      mapping={'s': sample_ID, 'kw': sample_service, 'e': repr(e)},
                      numline=row_nr, line=str(row))
            return
        self.processed_samples.append({sample_ID:sample_service})
        parsed.update({"DefaultResult": "Reading"})
        self._addRawResult(sample_ID, {keyword: parsed})
        return 0


    @staticmethod
    def is_sample(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        return True if brains else False


    @staticmethod
    def is_analysis_group_id(analysis_group_id):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_group_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        return True if brains else False


    @staticmethod
    def get_ar(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        try:
            return api.get_object(brains[0])
        except IndexError:
            pass


    @staticmethod
    def get_duplicate_or_qc(analysis_id,sample_service,):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, a) for a in brains)
        brains = [v for k, v in analyses.items() if k.startswith(sample_service)]
        if len(brains) < 1:
            msg = (" No analysis found matching Keyword {}".format(sample_service))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(sample_service))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    @staticmethod
    def get_reference_sample(reference_sample_id, kw):
        query = dict(
            portal_type="ReferenceSample", getId=reference_sample_id
        )
        brains = api.search(query, SENAITE_CATALOG)
        if len(brains) < 1:
            msg = ("No reference sample found matching Keyword {}".format(kw))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(kw))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    def get_reference_sample_analysis(self, reference_sample, kw):
        kw = kw
        brains = self.get_reference_sample_analyses(reference_sample)
        brains = [v for k, v in brains.items() if k.startswith(kw)]
        if len(brains) < 1:
            msg = " No analysis found matching Keyword {}".format(kw)
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(kw))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    @staticmethod
    def get_reference_sample_analyses(reference_sample):
        brains = reference_sample.getObject().getReferenceAnalyses()
        return dict((a.getKeyword(), a) for a in brains)


    def get_analysis(self, ar, kw):
        analyses = self.get_analyses(ar)
        analyses = [v for k, v in analyses.items() if k.startswith(kw)]
        if len(analyses) < 1:
            self.log(' No analysis found matching keyword {}'.format(kw))
            return None
        if len(analyses) > 1:
            self.warn('Multiple analyses found matching Keyword {}'.format(kw))
            return None
        return analyses[0]


    @staticmethod
    def get_analyses(ar):
        analyses = ar.getAnalyses()
        return dict((a.getKeyword, a) for a in analyses)


    @staticmethod
    def extract_relevant_data(lines):
        new_lines = []
        for row in lines:
            split_row = row.encode("ascii","ignore").split(",")
            if len(split_row) > 13:
                new_lines.append(','.join([str(elem) for elem in split_row]))
        return new_lines


    @staticmethod
    def try_utf8(data):
        """Returns a Unicode object on success, or None on failure"""
        try:
            return data.decode('utf-8')
        except UnicodeDecodeError:
            return None
    
    @staticmethod
    def try_utf16(data):
        """Returns a Unicode object on success, or None on failure"""
        try:
            return data.decode('utf-16')
        except UnicodeDecodeError:
            return None

    @staticmethod
    def parse_headerlines(reader):
        "To be implemented if necessary"
        return True


    @staticmethod
    def interim_map_sorter(row):
        interims = {}
        for k,v in row.items():
            sub = field_interim_map.get(k,'')
            if sub != '':
                interims[sub] = v
        return interims
    

    @staticmethod
    def data_cleaning(parsed):
        for k,v in parsed.items():
            #Sometimes a Factor value is not included in sheet
            if k == "Factor" and not v:
                parsed[k] = 1
            else:
                try:
                    parsed[k] = float(v)
                except (TypeError, ValueError):
                    parsed[k] = v
        return parsed


class dr3900import(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Hach DR3900"
    __file__ = abspath(__file__)

    def __init__(self, context):
        self.context = context
        self.request = None

    @staticmethod
    def Import(context, request):
        errors = []
        logs = []
        warns = []

        infile = request.form["instrument_results_file"]
        artoapply = request.form["artoapply"]
        override = request.form["results_override"]
        instrument = request.form.get("instrument", None)
        worksheet = request.form.get("worksheet", 0)

        ext = splitext(infile.filename.lower())[-1]
        if not hasattr(infile, "filename"):
            errors.append(_("No file selected"))

        parser = DR3900Parser(infile, worksheet=worksheet)

        if parser:

            status = ["sample_received", "attachment_due", "to_be_verified"]
            if artoapply == "received":
                status = ["sample_received"]
            elif artoapply == "received_tobeverified":
                status = ["sample_received", "attachment_due", "to_be_verified"]

            over = [False, False]
            if override == "nooverride":
                over = [False, False]
            elif override == "override":
                over = [True, False]
            elif override == "overrideempty":
                over = [True, True]

            importer = AnalysisResultsImporter(
                parser=parser,
                context=context,
                allowed_ar_states=status,
                allowed_analysis_states=None,
                override=over,
                instrument_uid=instrument,
            )

            try:
                importer.process()
                errors = importer.errors
                logs = importer.logs
                warns = importer.warns
            except Exception as e:
                errors.extend([repr(e), traceback.format_exc()])

        results = {"errors": errors, "log": logs, "warns": warns}

        return json.dumps(results)


class MyExport(BrowserView):

    def __innit__(self,context,request):
        self.context = context
        self.request = request
    

    def __call__(self,analyses):
        uc = api.get_tool('uid_catalog')
        instrument = self.context.getInstrument()
        filename = '{}-{}.csv'.format(
            self.context.getId(), instrument.Title())
        now = DateTime().strftime('%m/%d/%Y')

        layout = self.context.getLayout()
        tmprows = []
        parsed_analyses = {}
        headers = ["#Sample Number","#ID","#Date","#LIMS ID"]
        tmprows.append(headers)
        rows = []

        for indx,item in enumerate(layout):
            c_uid = item['container_uid']
            a_uid = item['analysis_uid']
            analysis = uc(UID=a_uid)[0].getObject() if a_uid else None
            container = uc(UID=c_uid)[0].getObject() if c_uid else None

            if item['type'] == 'a':
                analysis_id = container.id
            elif (item['type'] in 'bcd'):
                analysis_id = analysis.getReferenceAnalysesGroupID()
            if parsed_analyses.get(analysis_id):
                continue
            else:
                tmprows.append([indx+1,
                                analysis_id,
                                now,
                                ''])
                parsed_analyses[analysis_id] = 10

        rows = self.row_sorter(tmprows)
        result = self.dict_to_string(rows)

        setheader = self.request.RESPONSE.setHeader
        setheader('Content-Length', len(result))
        setheader('Content-Disposition', 'inline; filename=%s' % filename)
        setheader('Content-Type', 'text/csv')
        self.request.RESPONSE.write(result)
    

    @staticmethod
    def dict_to_string(rows):
        final_rows = ''
        interim_rows = []
        
        for row in rows:
            row = ','.join(str(item) for item in row)
            interim_rows.append(row)
        final_rows = '\r\n'.join(interim_rows)
        return final_rows

    
    @staticmethod
    def row_sorter(rows):
        for indx,row in enumerate(rows):
            if indx!= 0:
                row[0] = indx
        return rows


class dr3900export(object):
    implements(IInstrumentExportInterface)
    title = "Hach DR3900 Exporter"
    __file__ = abspath(__file__)  # noqa


    def __init__(self, context,request=None):
        self.context = context
        self.request = request


    def Export(self, context, request):
        return MyExport(context,request)