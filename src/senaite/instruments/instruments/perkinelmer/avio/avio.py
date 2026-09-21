# -*- coding: utf-8 -*-
"""Perkin Elmer Syngistix result import and SIFX worksheet export."""

import csv
import json
import re
import traceback
from decimal import Decimal
from decimal import InvalidOperation
from mimetypes import guess_type
from os.path import abspath
from os.path import splitext
from cStringIO import StringIO

from Products.CMFPlone.utils import safe_unicode
from bika.lims import api
from bika.lims.browser import BrowserView
from openpyxl import load_workbook
from zope.interface import implements

from senaite.instruments import senaiteMessageFactory as _
from senaite.core.catalog import ANALYSIS_CATALOG
from senaite.core.exportimport.instruments import (
    IInstrumentAutoImportInterface,
    IInstrumentExportInterface,
    IInstrumentImportInterface,
)
from senaite.core.exportimport.instruments.resultsimport import (
    AnalysisResultsImporter,
)
from senaite.instruments.instruments.pg.dv5000icp.dv5000 import (
    DV5000ICPParser,
)


WORKSHEET = "Conc. in Sample Units"

# Syngistix v5.5 requires the initial empty row and these fixed parameters.
SIFX_HEADER = (
    "",
    "[System Description]",
    "Description=Test LIMS import",
    "MaxNoOfSamples=250",
    "[Constant Parameters]",
    "BatchID=Batch 01",
    "VolumeUnits=Vol,mL,0.001000,,",
    "WeightUnits=Wt,g,1.000000,,",
    "[User Defined List]",
    "UserDefined1=",
    "UserDefined2=",
    "UserDefined3=",
    "UserDefined4=",
    "UserDefined5=",
    "[Variable Parameter List]",
    "NumberOfParameters=7",
    "Parameter1=SampleNo",
    "Parameter2=AutosamplerLocation",
    "Parameter3=SampleID",
    "Parameter4=InitialSampleWt",
    "Parameter5=SamplePrepVolume",
    "Parameter6=AliquotVolume",
    "Parameter7=DilutedToVolume",
    "[Variable Parameter Data]",
)


class SyngistixWorksheetExport(BrowserView):
    """Export one SIFX record per occupied worksheet sample position."""

    def __call__(self, analyses):
        rows = self.get_sample_rows()
        buffer = StringIO()
        buffer.write("\r\n".join(SIFX_HEADER) + "\r\n")
        buffer.write("NumberOfDataValues=%s\r\n" % len(rows))
        writer = csv.writer(buffer, lineterminator="\r\n")
        for number, row in enumerate(rows, 1):
            values = ["Data{0}={0}".format(number)] + row
            writer.writerow([unicode(safe_unicode(value)).encode("utf-8")
                             for value in values])
        result = buffer.getvalue()
        response = self.request.RESPONSE
        response.setHeader("Content-Type", "text/csv; charset=utf-8")
        response.setHeader("Content-Disposition",
                           'attachment; filename="%s.sifx"' %
                           self.context.getId())
        response.setHeader("Content-Length", len(result))
        response.write(result)

    def get_sample_rows(self):
        rows = []
        seen = set()
        layout = sorted(self.context.getLayoutView(),
                        key=lambda item: int(item["position"]))
        for item in layout:
            sample_type = item.get("type")
            if sample_type == "a":
                sample = api.get_object(item["container_uid"])
                sample_id = sample.getId()
                weight_analyses = sample.getAnalyses(
                    getKeyword="InitialSampleWt", full_objects=True)
            elif sample_type in ("b", "c", "d"):
                analysis = api.get_object(item["analysis_uid"])
                sample_id = analysis.getReferenceAnalysesGroupID()
                weight_analyses = api.search({
                    "portal_type": ["DuplicateAnalysis", "ReferenceAnalysis"],
                    "getReferenceAnalysesGroupID": sample_id,
                    "getKeyword": "InitialSampleWt",
                }, ANALYSIS_CATALOG)
            else:
                continue
            position = item["position"]
            key = (position, sample_id)
            if key in seen:
                continue
            seen.add(key)
            weight = ""
            if weight_analyses:
                weight = api.get_object(weight_analyses[0]).getResult()
                if weight is None:
                    weight = ""
            rows.append([position, sample_id, weight, "", "", ""])
        return rows


class syngistixexporter(object):
    implements(IInstrumentExportInterface)
    title = "Syngistix Worksheet Exporter"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context, request=None):
        self.context = context
        self.request = request

    def Export(self, context, request):
        return SyngistixWorksheetExport(context, request)


class AvioParser(DV5000ICPParser):
    """Read results exclusively from the Conc. in Sample Units sheet."""

    def __init__(self, infile, worksheet=WORKSHEET, encoding=None):
        # The sheet argument is accepted for importer API compatibility, but
        # this instrument must never read intensity or RSD sheets.
        super(AvioParser, self).__init__(
            infile, worksheet=WORKSHEET, encoding=encoding)

    def parse(self):
        ext = splitext(self.infile.filename.lower())[-1]
        if ext not in (".xlsx", ".xlsm"):
            self.err("Unsupported file format: %s" % ext)
            return -1

        try:
            workbook = load_workbook(filename=self.infile, data_only=True)
        except Exception:
            self.err("Cannot open workbook: %s" % self.infile.filename)
            return -1

        try:
            worksheet = workbook[WORKSHEET]
        except KeyError:
            self.err("Sheet not found in workbook: %s" % WORKSHEET)
            return -1

        # Syngistix repeats the header whenever the analyte set changes.
        headers = []
        for row_nr in range(1, worksheet.max_row + 1):
            sample_id = self.safe_value(
                worksheet.cell(row=row_nr, column=2).value)
            if sample_id == "Sample Id":
                headers = self.get_result_headers(worksheet, row_nr)
                continue
            if not sample_id or not headers:
                continue

            values = [
                self.safe_value(
                    worksheet.cell(row=row_nr, column=column).value)
                for column, header in headers
            ]
            self.headers = [header for column, header in headers]
            self.parse_result_row(sample_id, row_nr, values)

        return 1

    @staticmethod
    def get_result_headers(worksheet, row_nr):
        headers = []
        # Columns A-G contain run metadata. Results start in column H.
        for column in range(8, worksheet.max_column + 1):
            value = AvioParser.safe_value(
                worksheet.cell(row=row_nr, column=column).value)
            if value:
                headers.append((column, value))
        return headers

    @staticmethod
    def normalize_keyword(header):
        """Turn ``Fe-lb 238.863 (mg/L)`` into ``Fe238863``."""
        header = re.sub(r"\b(?:lb|cps|mg\s*/\s*l)\b", "", header,
                        flags=re.IGNORECASE)
        return re.sub(r"[^A-Za-z0-9]", "", header)

    def add_result(self, parsed, analyses, instrument_keyword, value,
                   sample_id, row_nr):
        """Keep Ca/Mg in ppm and convert other mg/L readings to ppb."""
        if instrument_keyword[:2] not in ("Ca", "Mg"):
            try:
                value = str(Decimal(value) * 1000)
            except InvalidOperation:
                # Preserve nonnumeric instrument flags for normal handling.
                pass
        return super(AvioParser, self).add_result(
            parsed, analyses, instrument_keyword, value, sample_id, row_nr)


class importer(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Perkin Elmer Syngistix Uranium"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context):
        self.context = context
        self.request = None

    @staticmethod
    def Import(context, request):
        errors = []
        logs = []
        warns = []

        infile = request.form["instrument_results_file"]
        if not hasattr(infile, "filename"):
            errors.append(_("No file selected"))

        artoapply = request.form["artoapply"]
        override = request.form["results_override"]
        instrument = request.form.get("instrument", None)
        parser = AvioParser(infile)

        status = ["sample_received", "attachment_due", "to_be_verified"]
        if artoapply == "received":
            status = ["sample_received"]

        over = [False, False]
        if override == "override":
            over = [True, False]
        elif override == "overrideempty":
            over = [True, True]

        results_importer = AnalysisResultsImporter(
            parser=parser,
            context=context,
            allowed_sample_states=status,
            allowed_analysis_states=None,
            override=over,
            instrument_uid=instrument,
        )
        try:
            results_importer.process()
            errors = results_importer.errors
            logs = results_importer.logs
            warns = results_importer.warns
        except Exception as error:
            errors.extend([repr(error), traceback.format_exc()])

        return json.dumps({"errors": errors, "log": logs, "warns": warns})

    def get_automatic_importer(self, instrument, parser, **kw):
        if getattr(parser, "_instrument", None) is None:
            parser._instrument = instrument
        self.parser = parser
        return self

    def get_automatic_parser(self, infile):
        return AvioParser(infile)

    def process(self):
        return self.Import(self.context, request=None)
