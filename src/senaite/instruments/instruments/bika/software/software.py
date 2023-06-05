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


class SampleNotFound(Exception):
    pass


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class SoftwareParser(InstrumentResultsFileParser):
    ar = None

    def __init__(self, infile, instrument, worksheet=None, encoding=None, delimiter=None):
        self.delimiter = delimiter if delimiter else ","
        self.encoding = encoding
        self.ar = None
        self.analyses = None
        self.worksheet = worksheet if worksheet else 0
        self.infile = infile
        self.csv_data = None
        self.csv_comments_data = None
        self.sample_id = None
        self.instrument_uid = instrument
        mimetype, encoding = guess_type(self.infile.filename)
        InstrumentResultsFileParser.__init__(self, infile, mimetype)

    def parse(self):
        order = []
        ext = splitext(self.infile.filename.lower())[-1]
        if ext == ".xlsx":
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
                    self.infile.seek(0)
                    self.csv_comments_data = importer(
                        infile=self.infile,
                        worksheet=1,
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
        comments_stub = FileStub(file=self.csv_comments_data, name=str(self.infile.filename))
        self.csv_data = FileUpload(stub)
        self.csv_comments_data = FileUpload(comments_stub)
        
        comments_data = self.csv_comments_data.read()
        data = self.csv_data.read()
        lines = self.decode_read_data(data,ext)
        comments_lines = self.decode_read_data(comments_data,ext)
        sample_comments = self.comments_parser(comments_lines)
        self.results_parser(lines,sample_comments)
        return 0
    
    def comments_parser(self,data):
        sample_comments = {}
        for row_num,row in enumerate(data):
            if row_num > 6:
                try:
                    sample_comments.update({row[0]:row[1]})
                except IndexError:
                    pass
        return sample_comments
    
    def results_parser(self,data,comments):
        for row_num,row in enumerate(data):
            if row_num == 4:
                sample_ids = row[1::2]
                sample_ids.pop(0)
                if not sample_ids[-1]:
                    sample_ids.pop(-1)
            if row_num > 7:
                if any(row[1:]): #checking if all row elements are non empty
                    clean_row = row[::]
                    clean_row.pop(1)
                    clean_row.pop(1)
                    self.parse_row(clean_row,row_num,sample_ids,comments)

    def decode_read_data(self,data,ext):
        decoded_data = self.try_utf8(data)
        if decoded_data:
            if ext == ".xlsx" or ext == ".xls":
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
        return ascii_lines

    def parse_row(self, row, row_nr,sample_ids,comments):
        sample_service = row.pop(0)
        for indx,sample_id in enumerate(sample_ids):
            result = row[2*indx]
            rl = row[(2*indx) + 1]
            if not result:
                continue
            try:
                if self.is_sample(sample_id):
                    ar = self.get_ar(sample_id)
                    analysis = self.get_analysis(ar,sample_service)
                    keyword = analysis.getKeyword
                elif self.is_analysis_group_id(sample_id): #Needs changing if will be used from key to title
                    analysis = self.get_duplicate_or_qc(sample_id,sample_service)
                    keyword = analysis.getKeyword
                else:
                    sample_reference = self.get_reference_sample(sample_id, sample_service) #Needs changing if will be used
                    analysis = self.get_reference_sample_analysis(sample_reference, sample_service)
                    keyword = analysis.getKeyword()
            except Exception as e:
                self.warn(msg="Error getting analysis for '${s}/${kw}': ${e}",
                        mapping={'s': sample_id, 'kw': sample_service, 'e': repr(e)},
                        numline=row_nr, line=str(row))
                return
            if rl:
                self.set_uncertainty(analysis,rl)
            # Allow manual editing of uncertainty must be ticked on analysis service
            if comments.get(sample_id):
                self.set_sample_remarks(ar,comments.pop(sample_id))
            result = self.result_detection_limit(result,analysis)
            parsed_results = {'Reading': result}
            parsed_results.update({"DefaultResult": "Reading"})
            self._addRawResult(sample_id, {keyword: parsed_results})
        return 0
    
    def result_detection_limit(self,result,analysis):
        if '<' not in result and '>' not in result:
            return result
        analysis_obj = analysis.getObject()
        operand = result[0]
        if operand == '<':
            ldl = analysis_obj.getLowerDetectionLimit()
            if ldl:
                result = float(ldl) - 1
            else:
                result = -1
        elif operand == '>':
            udl = analysis_obj.getUpperDetectionLimit()
            if udl:
                result = float(udl) + 1
            else:
                result = -1
        return str(result)
    
    def set_uncertainty(self,analysis,uncertainty_value):
        analysis_obj = analysis.getObject()
        analysis_obj.setUncertainty(uncertainty_value)

    def set_sample_remarks(self,sample,remark):
        instrument_obj = api.get_object_by_uid(self.instrument_uid)
        instrument_title = instrument_obj.Title() #use instrument name
        sample.setRemarks(api.safe_unicode(instrument_title+ ": " + remark))

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
        analyses = dict((a.Title, a) for a in brains)
        if len(brains) < 1:
            msg = (" No sample found with ID {}".format(analysis_id))
            raise AnalysisNotFound(msg)
        brains = [v for k, v in analyses.items() if k.startswith(sample_service)]
        if len(brains) < 1:
            msg = (" No analysis found matching Title {}".format(sample_service))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Title {}".format(sample_service))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    @staticmethod
    def get_reference_sample(reference_sample_id, kw):
        query = dict(
            portal_type="ReferenceSample", getId=reference_sample_id
        )#This method might not be needed
        brains = api.search(query, SENAITE_CATALOG)
        if len(brains) < 1:
            msg = ("No reference sample found with ID {}".format(reference_sample_id))
            raise AnalysisNotFound(msg)
        brains = [v for k, v in brains.items() if k == kw]
        if len(brains) < 1:
            msg = " No analysis found matching Keyword {}".format(kw)
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(kw))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    def get_reference_sample_analysis(self, reference_sample, title):
        title = title
        brains = self.get_reference_sample_analyses(reference_sample)
        if len(brains) < 1:
            msg = ("No sample found with ID {}".format(reference_sample)) 
            raise AnalysisNotFound(msg)
        brains = [v for k, v in brains.items() if k == title]
        if len(brains) < 1:
            msg = ("No analysis found matching Keyword {}".format(title))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(title))
            raise MultipleAnalysesFound(msg)
        return brains[0]


    @staticmethod
    def get_reference_sample_analyses(reference_sample):
        brains = reference_sample.getObject().getReferenceAnalyses()
        return dict((a.Title, a) for a in brains)


    def get_analysis(self, ar, title):
        analyses = self.get_analyses(ar)
        if len(analyses) < 1:
            msg = ' No sample found with ID {}'.format(ar)
            raise AnalysisNotFound(msg)
        analyses = [v for k, v in analyses.items() if k == title]
        if len(analyses) < 1:
            msg = ' No analysis found matching keyword {}'.format(title)
            raise AnalysisNotFound(msg)
        if len(analyses) > 1:
            msg = ' Multiple analyses found matching Keyword {}'.format(title)
            raise MultipleAnalysesFound(msg)
        return analyses[0]


    @staticmethod
    def get_analyses(ar):
        analyses = ar.getAnalyses()
        return dict((a.Title, a) for a in analyses)


    @staticmethod
    def extract_relevant_data(lines):
        new_lines = []
        for row in lines:
            split_row = row.encode("ascii","ignore").split(",")
            new_lines.append(split_row)
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


class softwareimport(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Bika Software"
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

        parser = SoftwareParser(infile, instrument,worksheet=worksheet)

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