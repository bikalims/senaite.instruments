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
        reader = csv.DictReader(ascii_lines)
        
        headers_parsed = self.parse_headerlines(reader)
        interim_fields = self.get_interim_fields()
        if headers_parsed:
            for row in reader:
                self.parse_row(row,reader.line_num,interim_fields)
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

    def parse_row(self, row, row_nr,interim_fields):
        parsed_strings = {}
        results,sample_id = self.get_results_values(row,row_nr)

        # parsed_strings = self.interim_map_sorter(results) #Still need to figure out what's happening with the interims
        # parsed = self.data_cleaning(parsed_strings)

        # regex = re.compile('[^a-zA-Z]')

        for sample_service in results.keys():
            if results.get(sample_service):
                try:
                    if self.is_sample(sample_id):
                        ar = self.get_ar(sample_id)
                        analysis = self.get_analysis(ar,sample_service)
                        keyword = analysis.getKeyword
                        if interim_fields.get(keyword):
                            pass #do something
                    elif self.is_analysis_group_id(sample_id):
                        analysis = self.get_duplicate_or_qc(sample_id,sample_service)
                        keyword = analysis.getKeyword
                        if interim_fields.get(keyword):
                            pass #do something
                    else:
                        sample_reference = self.get_reference_sample(sample_id, sample_service)
                        analysis = self.get_reference_sample_analysis(sample_reference, sample_service)
                        keyword = analysis.getKeyword()
                        if interim_fields.get(keyword):
                            pass #do something
                except Exception as e:
                    self.warn(msg="Error getting analysis for '${s}/${kw}': ${e}",
                            mapping={'s': sample_id, 'kw': sample_service, 'e': repr(e)},
                            numline=row_nr)
                    continue
            else:
                continue
            successfully_parsed = {}
            successfully_parsed[sample_service] = results.get(sample_service)
            successfully_parsed.update({"DefaultResult": sample_service})
            self._addRawResult(sample_id, {keyword: successfully_parsed})
        return 0

    @staticmethod
    def get_results_values(row,row_nr):
        barcode_ct = row.get('barcode_ct') #sample ID for other sample types
        barcode_boiler = row.get('barcode_boiler') #Sample ID for boiler water
        if barcode_ct:
            results = {} #ignore W and X
            sample_id = barcode_ct
            analysis_service_name_maybe = row.get("tower_system_name") #N WHat is this????????????????
            results["ControllerConductivity"] = row.get("controller_conductivity") #O
            results["field_conductivity_"] = row.get("field_conductivity_") #P 
            results["CalPerc"] = row.get("calibration_percent") #Q
            results["ControllerpH"] = row.get("controller_ph") #R 
            results["field_ph_"] = row.get("field_ph_") #S 
            results["ControllerORP"] = row.get("controller_orp") #T
            results["FAH"] = row.get("free_available_halogen") #U
            results["TAH"] = row.get("total_available_halogen") #V
            results["controller_trasar_value"] = row.get("controller_tracer_value") #Y
            results["ControllerPTSA"] = row.get("controller_ptsa_value") #Z
            results["ControllerTrasar"] = row.get("controller_trasar_value") #AA
            results["ControllerTAG"] = row.get("controller_tag_value") #AB
            results["controller_ptsa_value"] = row.get("controller_pyxis_value") #AC
        elif barcode_boiler:
            analysis_service_name_maybe = row.get("tower_system_name") #remove this?
            sample_id = barcode_boilder
            results = {}
            results["SulfiteasSO2"] = row.get("sulfite_test_result") #CC
            results["boiler_field_conductivity"] = row.get("boiler_controller_conductivity") #CD Still to be added to Bika
            results["BFC"] = row.get("boiler_field_conductivity") #CE
        else:
            analysis_service_name_maybe = row.get("tower_system_name") #remove this?
            no_id_results = {}
            no_id_results["SulfiteasSO2"] = row.get("sulfite_test_result") #CC
            no_id_results["boiler_field_conductivity"] = row.get("boiler_controller_conductivity") #CD Still to be added to Bika
            no_id_results["BFC"] = row.get("boiler_field_conductivity")
            analysis_service_name_maybe = row.get("tower_system_name") #N remove maybe?
            no_id_results["ControllerConductivity"] = row.get("controller_conductivity") #O
            no_id_results["field_conductivity_"] = row.get("field_conductivity_") #P 
            no_id_results["CalPerc"] = row.get("calibration_percent") #Q
            no_id_results["ControllerpH"] = row.get("controller_ph") #R 
            no_id_results["field_ph_"] = row.get("field_ph_") #S 
            no_id_results["ControllerORP"] = row.get("controller_orp") #T
            no_id_results["FAH"] = row.get("free_available_halogen") #U
            no_id_results["TAH"] = row.get("total_available_halogen") #V
            no_id_results["controller_trasar_value"] = row.get("controller_tracer_value") #Y
            no_id_results["ControllerPTSA"] = row.get("controller_ptsa_value") #Z
            no_id_results["ControllerTrasar"] = row.get("controller_trasar_value") #AA
            no_id_results["ControllerTAG"] = row.get("controller_tag_value") #AB
            no_id_results["controller_ptsa_value"] = row.get("controller_pyxis_value") #AC
            if any(no_id_results.values()):
                msg = ("No Sample ID was found for results on row {0} and {1}. Please capture results manually".format(row_nr,analysis_service_name_maybe))
                raise SampleNotFound(msg)
        return results,sample_id

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
    def get_interim_fields():
        bsc = api.get_tool("senaite_catalog_setup")
        query = {
            "portal_type": "AnalysisService",
            "is_active": True,
            "sort_on": "sortable_title",
        }
        services = bsc(query,)
        all_services = {}
        for y in services:
            service_obj = y.getObject()
            all_services[service_obj.getKeyword()] = service_obj.getInterimFields()
        return all_services

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
        artoapply = request.form["artoapply"]
        override = request.form["results_override"]
        instrument = request.form.get("instrument", None)
        worksheet = request.form.get("worksheet", 0)

        ext = splitext(infile.filename.lower())[-1]
        if not hasattr(infile, "filename"):
            errors.append(_("No file selected"))

        parser = FulcrumAppParser(infile, worksheet=worksheet)

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