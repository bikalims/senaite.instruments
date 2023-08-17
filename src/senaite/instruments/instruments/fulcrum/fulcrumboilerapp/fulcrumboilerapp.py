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
from re import subn
from zope.interface import implements
from zope.publisher.browser import FileUpload


class SampleNotFound(Exception):
    pass


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class FulcrumBoilerAppParser(InstrumentResultsFileParser):
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
        if ext == ".xlsx":  # check in flameatomic also
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
                    self.err(
                        "Sheet not found in workbook: %s" % self.worksheet)
                    return -1
                except Exception as e:
                    pass
            else:
                self.warn("Can't parse input file as XLS, XLSX, or CSV.")
                return -1

        stub = FileStub(file=self.csv_data, name=str(self.infile.filename))
        self.csv_data = FileUpload(stub)
        lines = self.csv_data.readlines()
        reader = csv.DictReader(lines)
        for row in reader:
            results = self.get_result_values(row, reader.line_num)
            if results:
                self.parse_row(results, reader.line_num)
        return 0

    def get_result_values(self, results, row_num):
        condenser = [
                "controller_conductivity",
                "field_conductivity_",
                "calibration_percent",
                "controller_ph",
                "field_ph_",
                "controller_orp",
                "free_available_halogen",
                "total_available_halogen",
                "controller_tracer_type",
                "controller_tracer_type_other",
                "controller_tracer_value",
                "controller_ptsa_value",
                "controller_trasar_value",
                "controller_tag_value",
                ]
        boiler = [
                "sulfite_test_result",
                "boiler_controller_conductivity",
                "boiler_field_conductivity",
                ]
        nb_results = {}

        if results.get("barcode_ct"):
            sample_id = results.get("barcode_ct")
            nb_results["sample_id"] = subn(r'[^\w\d\-_]*', '', sample_id)[0]
            for kw in condenser:
                nb_results[subn(r'[^\w\d]*', '', kw)[0]] = results.get(kw)

        elif results.get("barcode_boiler"):
            sample_id = results.get("barcode_boiler")
            nb_results["sample_id"] = subn(r'[^\w\d\-_]*', '', sample_id)[0]
            for kw in boiler:
                nb_results[subn(r'[^\w\d]*', '', kw)[0]] = results.get(kw)

        else:
            self.warn(
                msg="No Sample ID was found for results on row '${r}'."
                    " Please capture results manually",
                mapping={'r': str(row_num)})
        return nb_results

    def parse_row(self, row, row_num):
        sample_id = row.get("sample_id")
        row.pop("sample_id")
        interim_keywords = self.get_interim_fields(sample_id)

        for sample_service in row.keys():
            reading = row.get(sample_service)
            if reading:
                if {sample_id: sample_service} in self.processed_samples:
                    msg = ("Multiple results for Sample '{}' "
                           "with analysis service '{}'"
                           " found. Not imported".format(
                                                    sample_id,
                                                    sample_service))
                    raise MultipleAnalysesFound(msg)
                    continue
                should_break = self.try_getting_analysis(
                                sample_id,
                                sample_service,
                                reading,
                                row_num,
                                interim_keywords
                )
                if should_break:  # sample doesn't exist
                    break
            else:
                self.warn(
                    msg="No results found for '${id}'/'${service}'",
                    mapping={"id": sample_id, "service": sample_service},
                    numline=str(row_num),
                )
                continue
        return

    def try_getting_analysis(
            self, sample_id, sample_service,
            reading, row_num, interim_keywords):

        portal_type = self.get_portal_type(sample_id)
        analysis = ""  # will be updated in try block
        keyword = ""  # will be updated in try block
        try:
            if portal_type == "AnalysisRequest":
                ar = self.get_ar(sample_id)
                analysis = self.get_analysis(ar, sample_service)
                keyword = analysis.getKeyword
            elif portal_type in ["DuplicateAnalysis", "ReferenceAnalysis"]:
                analysis = self.get_duplicate_or_qc(sample_id, sample_service)
                keyword = analysis.getKeyword
            elif portal_type == "ReferenceSample":
                sample_reference = self.get_reference_sample(
                                    sample_id, sample_service)
                analysis = self.get_reference_sample_analysis(
                                    sample_reference, sample_service)
                keyword = analysis.getKeyword()
            else:
                self.warn(
                    msg="No Sample '${s}' found for results on row '${r}'."
                    " Results have not been imported",
                    mapping={'s': sample_id, 'r': str(row_num)}
                )
                return "break"

        except Exception as e:
            if not analysis:
                keyword = self.process_interims(
                                interim_keywords,
                                sample_service,
                                sample_id,
                                reading
                            )
                if not keyword:
                    self.warn(
                            msg="Error getting analysis for"
                                " '${s}/${kw}': ${e}",
                            mapping={
                                's': sample_id,
                                'kw': sample_service,
                                'e': repr(e)},
                            numline=str(row_num))
                    return
                else:
                    return
        if keyword:
            self.parse_results(reading, keyword, sample_id)
        return

    def get_portal_type(self, sample_id):
        portal_type = None
        if self.is_sample(sample_id):
            ar = self.get_ar(sample_id)
            self.ar = ar
            self.analyses = self.get_analyses(ar)
            portal_type = ar.portal_type
        elif self.is_analysis_group_id(sample_id):
            portal_type = "DuplicateAnalysis"
        elif self.is_reference_sample(sample_id):
            portal_type = "ReferenceSample"
        return portal_type

    @staticmethod
    def is_reference_sample(reference_sample_id):
        query = dict(portal_type="ReferenceSample", getId=reference_sample_id)
        brains = api.search(query, SENAITE_CATALOG)
        return True if brains else False

    @staticmethod
    def get_interim_fields(sample_id):
        bc = api.get_tool(CATALOG_ANALYSIS_REQUEST_LISTING)
        ar = bc(portal_type='AnalysisRequest', id=sample_id)
        if len(ar) == 0:
            ar = bc(
                portal_type='AnalysisRequest', getClientSampleID=sample_id)
        if len(ar) == 1:
            obj = ar[0].getObject()
            analyses = obj.getAnalyses(full_objects=True)
            keywords = {}
            for analysis_service in analyses:
                for field in analysis_service.getInterimFields():
                    interim_kw = field.get("keyword")
                    as_kw = analysis_service.Keyword
                    if interim_kw in keywords.keys():
                        keywords[interim_kw].append(as_kw)
                    else:
                        keywords[interim_kw] = [as_kw]
            return keywords
        return {}

    def process_interims(self, interims, kw, sample_id, result):
        as_kw = interims.get(kw)
        if as_kw:
            if len(as_kw) > 1:
                self.warn("Duplicate keyword {0} found for Sample {1} and"
                          " their results are not imported".format(
                                                kw, sample_id))
                return "Duplicate"
            else:
                self.processed_samples.append({sample_id: kw})
                self.parse_interims(result, as_kw[0], kw, sample_id)
                return as_kw[0]
        else:
            return

    def parse_interims(self, result, as_kw, interim_kw, sample_id):
        parsed = {}
        parsed[interim_kw] = result
        parsed.update({"DefaultResult": interim_kw})
        self._addRawResult(sample_id, {as_kw: parsed})

    def parse_results(self, result, keyword, sample_id):
        self.processed_samples.append({sample_id: keyword})
        parsed = {}
        parsed[keyword] = result
        parsed.update({"DefaultResult": keyword})
        self._addRawResult(sample_id, {keyword: parsed})

    @staticmethod
    def is_sample(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, CATALOG_ANALYSIS_REQUEST_LISTING)
        return True if brains else False

    @staticmethod
    def is_analysis_group_id(analysis_group_id):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types,
            getReferenceAnalysesGroupID=analysis_group_id
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
    def get_duplicate_or_qc(analysis_id, sample_service,):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(
            portal_type=portal_types, getReferenceAnalysesGroupID=analysis_id
        )
        brains = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, a) for a in brains)
        if len(brains) < 1:
            msg = (" No sample found with ID {}".format(analysis_id))
            raise AnalysisNotFound(msg)
        brains = [
                v for k,
                v in analyses.items() if k.startswith(sample_service)]
        if len(brains) < 1:
            msg = (" No analysis found matching Keyword {}".format(
                                                        sample_service))
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            msg = ("Multiple brains found matching Keyword {}".format(
                                                        sample_service))
            raise MultipleAnalysesFound(msg)
        return brains[0]

    @staticmethod
    def get_reference_sample(reference_sample_id, kw):
        query = dict(
            portal_type="ReferenceSample", getId=reference_sample_id
        )
        brains = api.search(query, SENAITE_CATALOG)
        if len(brains) < 1:
            msg = ("No reference sample found with ID {}".format(
                                                        reference_sample_id))
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


class fulcrumboilerappimport(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Fulcrum Boiler App"
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

        parser = FulcrumBoilerAppParser(infile, worksheet=worksheet)
        if parser:
            status = ["sample_received", "sample_due", "to_be_sampled"]
            over = [False, False]
            if override == "nooverride":
                over = [False, False]
            elif override == "override":
                over = [True, False]
            elif override == "overrideempty":
                over = [True, True]
            analysis_states = [
                    "unassigned", "assigned", "registered", "to_be_verified"]
            # ["unassigned","assigned","to_be_verified","rejected",
            # "retracted","verified","published","registered"] all of them

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
