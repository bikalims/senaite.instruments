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


class FulcrumAppParser(InstrumentResultsFileParser):
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
        
        lines_with_parentheses = self.use_correct_unicode_tranformation_format(data,ext)
        lines = [i.replace('"','') for i in lines_with_parentheses]
        ascii_lines = self.extract_relevant_data(lines)
        headers = ascii_lines[0]
        for row_nr,row in enumerate(ascii_lines):
            if row_nr!=0 and len(row) > 1:
                self.parse_row(row,row_nr,headers)
        return 0
    
    def use_correct_unicode_tranformation_format(self,data,ext):
        decoded_data = self.try_utf8(data)
        if decoded_data:
                lines_with_parentheses = decoded_data.split("\n")
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
        return lines_with_parentheses

    def parse_row(self, row, row_nr,headers):
        results,sample_id = self.get_results_values(row,row_nr,headers)
        if sample_id == 'None':
            return
        interim_keywords = self.get_interim_fields(sample_id)        

        for sample_service in results.keys():
            if results.get(sample_service):
                try:
                    if self.is_sample(sample_id):
                        ar = self.get_ar(sample_id)
                        analysis = self.get_analysis(ar,sample_service)
                        keyword = analysis.getKeyword
                    elif self.is_analysis_group_id(sample_id):
                        analysis = self.get_duplicate_or_qc(sample_id,sample_service)
                        keyword = analysis.getKeyword
                    else:
                        sample_reference = self.get_reference_sample(sample_id, sample_service)
                        analysis = self.get_reference_sample_analysis(sample_reference, sample_service)
                        keyword = analysis.getKeyword()
                except Exception as e:
                    self.warn(msg="Error getting analysis for '${s}/${kw}': ${e}",
                            mapping={'s': sample_id, 'kw': sample_service, 'e': repr(e)},
                            numline=row_nr)
                    continue
            else:
                if sample_service in interim_keywords.keys(): # for interims   # 'Zinc': [] , #{'Reading': 4.0, 'Factor': 1.0}
                    interimKeyword = interim_keywords.get(sample_service)
                    successfully_parsed = {}
                    successfully_parsed[interimKeyword] = results.get(sample_service)
                    successfully_parsed.update({"DefaultResult": interimKeyword})
                    self._addRawResult(sample_id, {sample_service: successfully_parsed})
                    continue
                else:
                    continue
            successfully_parsed = {}
            successfully_parsed[sample_service] = results.get(sample_service)
            successfully_parsed.update({"DefaultResult": sample_service})
            self._addRawResult(sample_id, {keyword: successfully_parsed})
        return 0

    def get_results_values(self,row,row_nr,headers):
        barcode_ct = row[12] #sample ID for other sample types M
        barcode_boiler = row[77] #Sample ID for boiler water BZ
        if barcode_ct:
            results = {} #ignore W and X,Y,AC
            sample_id = barcode_ct
            results[headers[13]] = row[13] #N ?
            results[headers[14]] = row[14] #O
            results[headers[15]] = row[15] #P 
            results[headers[16]] = row[16] #Q
            results[headers[17]] = row[17] #R 
            results[headers[18]] = row[18] #S 
            results[headers[19]] = row[19] #T
            results[headers[20]] = row[20] #U
            results[headers[21]] = row[21] #V
            results[headers[25]] = row[25] #Z
            results[headers[26]] = row[26] #AA
            results[headers[27]] = row[27] #AB
        elif barcode_boiler:
            sample_id = barcode_boiler
            results = {}
            results[headers[80]] = row[80] #CC
            results[headers[81]] = row[81] #CD Still to be added to Bika
            results[headers[82]] = row[82] #CE
        else:
            #regular sample results
            no_id_results = {}
            no_id_results[headers[13]] = row[13] #N ?
            no_id_results[headers[14]] = row[14] #O
            no_id_results[headers[15]] = row[15] #P 
            no_id_results[headers[16]] = row[16] #Q
            no_id_results[headers[17]] = row[17] #R 
            no_id_results[headers[18]] = row[18] #S 
            no_id_results[headers[19]] = row[19] #T
            no_id_results[headers[20]] = row[20] #U
            no_id_results[headers[21]] = row[21] #V
            no_id_results[headers[24]] = row[24] #Y
            no_id_results[headers[25]] = row[25] #Z
            no_id_results[headers[26]] = row[26] #AA
            no_id_results[headers[27]] = row[27] #AB
            no_id_results[headers[28]] = row[28] #AC
            #boiler sample results
            no_id_results[headers[80]] = row[80] #CC
            no_id_results[headers[81]] = row[81] #CD Still to be added to Bika
            no_id_results[headers[82]] = row[82] #CE
            if any(no_id_results.values()):
                self.warn(msg="No Sample ID was found for results on row '${r}'. Please capture results manually",
                    mapping={'r': row_nr})
            results = no_id_results
            sample_id = 'None'
        return results,sample_id

    @staticmethod
    def get_interim_fields(sample_id):
        bc = api.get_tool(CATALOG_ANALYSIS_REQUEST_LISTING)
        ar = bc(portal_type='AnalysisRequest', id=sample_id)
        if len(ar) == 0:
            ar = bc(portal_type='AnalysisRequest', getClientSampleID=sample_id)
        if len(ar) == 1:
            obj = ar[0].getObject()
            analyses = obj.getAnalyses(full_objects=True)
            services_with_interims = {}
            keywords = {}
            for analysis_service in analyses:
                if analysis_service.getInterimFields():
                    for field in analysis_service.getInterimFields():
                        keywords[analysis_service.getKeyword()] = field.get('keyword') 
            return keywords
        return {}

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
        if len(brains) < 1:
            msg = (" No sample found with ID {}".format(analysis_id))
            raise AnalysisNotFound(msg)
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

    def get_reference_sample_analysis(self, reference_sample, kw):
        kw = kw
        brains = self.get_reference_sample_analyses(reference_sample)
        if len(brains) < 1:
            msg = ("No sample found with ID {}".format(reference_sample)) 
            raise AnalysisNotFound(msg)
        brains = [v for k, v in brains.items() if k == kw]
        if len(brains) < 1:
            msg = ("No analysis found matching Keyword {}".format(kw))
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
        if len(analyses) < 1:
            msg = ' No sample found with ID {}'.format(ar)
            raise AnalysisNotFound(msg)
        analyses = [v for k, v in analyses.items() if k == kw]
        if len(analyses) < 1:
            msg = ' No analysis found matching keyword {}'.format(kw)
            raise AnalysisNotFound(msg)
        if len(analyses) > 1:
            msg = ' Multiple analyses found matching Keyword {}'.format(kw)
            raise MultipleAnalysesFound(msg)
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


class fulcrumappimport(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "FulcrumApp"
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
        override = request.form["results_override"]
        instrument = request.form.get("instrument", None)
        worksheet = request.form.get("worksheet", 0)

        ext = splitext(infile.filename.lower())[-1]
        if not hasattr(infile, "filename"):
            errors.append(_("No file selected"))

        parser = FulcrumAppParser(infile, worksheet=worksheet)
        if parser:
            status = ["sample_received", "sample_due", "to_be_sampled"]
            over = [False, False]
            if override == "nooverride":
                over = [False, False]
            elif override == "override":
                over = [True, False]
            elif override == "overrideempty":
                over = [True, True]
            analysis_states = ["unassigned","assigned","to_be_verified","rejected","retracted","verified","published","registered"] #all of them

            importer = AnalysisResultsImporter(
                parser=parser,
                context=context,
                allowed_ar_states=status,
                allowed_analysis_states=analysis_states,
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