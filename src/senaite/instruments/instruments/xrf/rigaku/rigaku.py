# -*- coding: utf-8 -*-
"""Importer for Rigaku XRF CSV result files."""

import csv
import json
import traceback
from cStringIO import StringIO
from mimetypes import guess_type
from os.path import abspath
from os.path import splitext

from zope.interface import implements

from senaite.core.exportimport.instruments import IInstrumentAutoImportInterface
from senaite.core.exportimport.instruments import IInstrumentImportInterface
from senaite.core.exportimport.instruments.resultsimport import AnalysisResultsImporter
from senaite.core.exportimport.instruments.resultsimport import InstrumentResultsFileParser
from senaite.instruments import senaiteMessageFactory as _
from senaite.instruments.instruments.xrf.axios.axios import AxiosXRFParser


class RigakuXRFParser(AxiosXRFParser):
    """Read sample IDs from column B and first-block results from column E."""

    def __init__(self, infile, encoding=None):
        self.infile = infile
        self.encoding = encoding or "utf-8"
        mimetype, unused = guess_type(infile.filename)
        InstrumentResultsFileParser.__init__(self, infile, mimetype)

    def parse(self):
        extension = splitext(self.infile.filename.lower())[-1]
        if extension != ".csv":
            self.err("Unsupported file format: %s" % extension)
            return -1
        try:
            self.infile.seek(0)
            data = self.infile.read()
            if isinstance(data, unicode):
                data = data.encode(self.encoding)
            rows = list(csv.reader(StringIO(data)))
        except Exception:
            self.err("Unable to read '{}':\n{}".format(
                self.infile.filename, traceback.format_exc()))
            return -1

        instrument_keyword = self.normalize_keyword(
            self.cell(rows[0], 4) if rows else "")
        if not instrument_keyword:
            self.err("Could not find an analysis header in column E")
            return -1

        # Row two contains units. Stop at the first blank sample ID so that
        # summary rows and any later analysis blocks are never imported.
        for row_nr, row in enumerate(rows[2:], 3):
            sample_id = self.cell(row, 1)
            if not sample_id:
                break
            value = self.cell(row, 4)
            if not value:
                continue
            if not self.is_number(value):
                self.warn(msg="Invalid result '${result}' for '${sample_id}'",
                          mapping={"result": value, "sample_id": sample_id},
                          numline=row_nr)
                continue
            analyses = self.get_analyses(sample_id)
            if not analyses:
                self.warn(msg="Sample or QC ID '${sample_id}' was not found",
                          mapping={"sample_id": sample_id}, numline=row_nr)
                continue
            parsed = {}
            self.add_result(parsed, analyses, instrument_keyword, value,
                            row, {}, sample_id, row_nr)
            if parsed:
                self._addRawResult(sample_id, parsed)
        return 1


class importer(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "Rigaku XRF"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context):
        self.context = context
        self.parser = None

    @staticmethod
    def Import(context, request):
        errors, logs, warns = [], [], []
        infile = request.form["instrument_results_file"]
        if not hasattr(infile, "filename"):
            return json.dumps({"errors": [_('No file selected')],
                               "log": logs, "warns": warns})
        parser = RigakuXRFParser(infile)
        states = ["sample_received", "attachment_due", "to_be_verified"]
        if request.form.get("artoapply") == "received":
            states = ["sample_received"]
        override = request.form.get("results_override", "overrideempty")
        over = {"override": [True, False],
                "overrideempty": [True, True]}.get(override, [False, False])
        results_importer = AnalysisResultsImporter(
            parser=parser, context=context, allowed_sample_states=states,
            allowed_analysis_states=None, override=over,
            instrument_uid=request.form.get("instrument"))
        try:
            results_importer.process()
            errors.extend(results_importer.errors)
            logs.extend(results_importer.logs)
            warns.extend(results_importer.warns)
        except Exception:
            errors.append(traceback.format_exc())
        return json.dumps({"errors": errors, "log": logs, "warns": warns})

    def get_automatic_parser(self, infile):
        return RigakuXRFParser(infile)

    def get_automatic_importer(self, instrument, parser, **kwargs):
        self.parser = parser
        return self

    def process(self):
        request = type("Request", (object,), {})()
        request.form = dict(instrument_results_file=self.parser.infile,
                            artoapply="received_tobeverified",
                            results_override="overrideempty", instrument=None)
        return self.Import(self.context, request)

