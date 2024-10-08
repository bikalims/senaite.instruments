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
# Copyright 2018-2024 by it's authors.
# Some rights reserved, see README and LICENSE.

import json
import traceback
from os.path import abspath
from re import subn
from zope.interface import implements

from bika.lims import api
from senaite.core.catalog import SAMPLE_CATALOG
from senaite.core.catalog import ANALYSIS_CATALOG, SENAITE_CATALOG
from senaite.instruments import senaiteMessageFactory as _
from senaite.core.exportimport.instruments import (
    IInstrumentAutoImportInterface,
    IInstrumentImportInterface,
)
from senaite.core.exportimport.instruments.resultsimport import (
    AnalysisResultsImporter,
    InstrumentCSVResultsFileParser,
)


class MultipleAnalysesFound(Exception):
    pass


class AnalysisNotFound(Exception):
    pass


class importer(object):
    implements(IInstrumentImportInterface, IInstrumentAutoImportInterface)
    title = "XRF"
    __file__ = abspath(__file__)  # noqa

    def __init__(self, context):
        self.context = context
        self.request = None

    @staticmethod
    def Import(context, request):
        """ Read XRF results
        """
        form = request.form
        # TODO form['file'] sometimes returns a list
        infile = form['instrument_results_file'][0] if \
            isinstance(form['instrument_results_file'], list) \
            else form['instrument_results_file']
        override = form['results_override']
        artoapply = form['artoapply']
        instrument = form.get('instrument', None)
        errors = []
        logs = []

        parser = None
        if not hasattr(infile, 'filename'):
            errors.append(_("No file selected"))
        infile.filename = infile.filename.replace("txt", "tsv")

        parser = XRFTXTParser2(infile)

        if parser:
            # Load the importer
            status = ["sample_received", "attachment_due", "to_be_verified"]
            if artoapply == "received":
                status = ["sample_received"]
            elif artoapply == "received_tobeverified":
                status = ["sample_received",
                          "attachment_due",
                          "to_be_verified"]

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
                                   allowed_sample_states=status,
                                   allowed_analysis_states=None,
                                   override=over,
                                   instrument_uid=instrument)
            tbex = ''
            try:
                importer.process()
            except Exception:
                tbex = traceback.format_exc()
            errors = importer.errors
            logs = importer.logs
            warns = importer.warns
            if tbex:
                errors.append(tbex)

        results = {'errors': errors, 'log': logs, 'warns': warns}
        return json.dumps(results)


class XRFTXTParser2(InstrumentCSVResultsFileParser):
    HEADERTABLE = []
    HEADERTABLE_DATA = []
    COMMAS = '\t '

    def __init__(self, csv):
        InstrumentCSVResultsFileParser.__init__(self, csv)
        self._end_header = False
        self._resultsheader = []
        self._numline = 0

    def _parseline(self, line):
        if self._end_header is False:
            return self.parse_headerline(line)
        elif self._end_header and not self._resultsheader:
            return self.parse_result_headerline(line)
        else:
            return self.parse_resultsline(line)

    def parse_headerline(self, line):
        """
        """
        if self._end_header is True:
            # Header already processed
            return 0

        splitted = [token.strip() for token in line.split(self.COMMAS)]

        # [Header]
        self._header = {item: '' for item in splitted}
        self.HEADERTABLE = splitted
        self._end_header = True
        return 0

    def parse_result_headerline(self, line):
        """ Parses quantitation result lines
            Please see samples/GC-MS output.txt
            [MS Quantitative Results] section
        """

        splitted = [token.strip() for token in line.split('\t')]
        if len(splitted[8:])/2 == len(self.HEADERTABLE):
            self.HEADERTABLE_DATA = splitted
            self._resultsheader = splitted
            return 0

    def parse_resultsline(self, line):
        splitted = [token.strip() for token in line.split('\t')]
        sample_id = splitted[1]
        line_num = splitted[0]
        results = splitted[8:]
        values = [results[i] for i in range(0, len(results), 2)]
        data = dict(zip(self.HEADERTABLE, values))

        portal_type = self.get_portal_type(sample_id)
        data["sample_id"] = sample_id

        for keyword in data:
            if keyword == "sample_id":
                continue
            if portal_type == "AnalysisRequest":
                self.parse_ar_row(keyword, line_num, data)

            elif portal_type in ["DuplicateAnalysis", "ReferenceAnalysis"]:
                self.parse_duplicate_row(keyword, line_num, data)
            elif portal_type == "ReferenceSample":
                self.parse_reference_sample_row(keyword, line_num, data)
            else:
                self.warn(
                    msg="No results found for '${sample_id}'",
                    mapping={"sample_id": sample_id},
                    numline=str(line_num),
                )
        return 0

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

    def parse_row(self, row_nr, parsed, keyword):
        parsed.update({"DefaultResult": keyword})
        self._addRawResult(parsed.get("sample_id"), {keyword: parsed})
        return 0

    def parse_ar_row(self, keyword, row_nr, row):
        sample_id = row.get("sample_id")
        ar = self.get_ar(sample_id)
        items = row.items()
        parsed = {subn(r'[^\w\d\-_]*', '', k)[0]: v for k, v in items if k}

        try:
            analysis = self.get_analysis(ar, keyword)
            if not analysis:
                return 0
        except Exception:
            self.warn(
                msg="Error getting analysis for '${kw}': ${sample_id}",
                mapping={"kw": keyword, "sample_id": sample_id},
                numline=row_nr,
            )
            return
        return self.parse_row(row_nr, parsed, keyword)

    def parse_duplicate_row(self, keyword, row_nr, row):
        items = row.items()
        parsed = {subn(r'[^\w\d\-_]*', '', k)[0]: v for k, v in items if k}
        sample_id = row.get("sample_id")

        try:
            if not self.getDuplicateKeyord(sample_id, keyword):
                return 0
        except Exception:
            self.warn(
                msg="Error getting analysis for '${kw}': ${sample_id}",
                mapping={"kw": keyword, "sample_id": sample_id},
                numline=row_nr,
            )
            return
        return self.parse_row(row_nr, parsed, keyword)

    def getDuplicateKeyord(self, sample_id, kw):
        analysis = self.get_duplicate_or_qc_analysis(sample_id, kw)
        return analysis.getKeyword

    def parse_reference_sample_row(self, keyword, row_nr, row):
        items = row.items()
        parsed = {subn(r'[^\w\d\-_]*', '', k)[0]: v for k, v in items if k}
        sample_id = row.get("sample_id")

        try:
            if not self.getReferenceSampleKeyword(sample_id, keyword):
                return 0
        except Exception:
            self.warn(
                msg="Error getting analysis for '${kw}': ${sample_id}",
                mapping={"kw": keyword, "sample_id": sample_id},
                numline=row_nr,
            )
            return
        return self.parse_row(row_nr, parsed, keyword)

    def getReferenceSampleKeyword(self, sample_id, kw):
        sample_reference = self.get_reference_sample(sample_id, kw)
        analysis = self.get_reference_sample_analysis(sample_reference, kw)
        return analysis.getKeyword()

    def get_reference_sample_analysis(self, reference_sample, kw):
        kw = kw
        brains = self.get_reference_sample_analyses(reference_sample)
        brains = [v for k, v in brains.items() if k == kw]
        if len(brains) < 1:
            lmsg = "No analysis found for sample {} matching Keyword {}"
            msg = lmsg.format(reference_sample, kw)
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            lmsg = "Multiple objects found for sample {} matching Keyword '{}'"
            msg = lmsg.format(reference_sample, kw)
            raise MultipleAnalysesFound(msg)
        return brains[0]

    @staticmethod
    def get_reference_sample_analyses(reference_sample):
        brains = reference_sample.getObject().getReferenceAnalyses()
        return dict((a.getKeyword(), a) for a in brains)

    @staticmethod
    def get_duplicate_or_qc_analysis(analysis_id, kw):
        portal_types = ["DuplicateAnalysis", "ReferenceAnalysis"]
        query = dict(portal_type=portal_types,
                     getReferenceAnalysesGroupID=analysis_id)
        brains = api.search(query, ANALYSIS_CATALOG)
        analyses = dict((a.getKeyword, a) for a in brains)
        brains = [v for k, v in analyses.items() if k == kw]
        if len(brains) < 1:
            lmsg = "No analysis found for sample {} matching Keyword {}"
            msg = lmsg.format(analysis_id, kw)
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            lmsg = "Multiple objects found for sample {} matching Keyword {}"
            msg = lmsg.format(analysis_id, kw)
            raise MultipleAnalysesFound(msg)
        return brains[0]

    @staticmethod
    def get_ar(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, SAMPLE_CATALOG)
        try:
            return api.get_object(brains[0])
        except IndexError:
            pass

    @staticmethod
    def is_sample(sample_id):
        query = dict(portal_type="AnalysisRequest", getId=sample_id)
        brains = api.search(query, SAMPLE_CATALOG)
        return True if brains else False

    @staticmethod
    def get_analyses(ar):
        analyses = ar.getAnalyses()
        return dict((a.getKeyword, a) for a in analyses)

    def get_analysis(self, ar, kw):
        analyses = self.get_analyses(ar)
        analyses = [v for k, v in analyses.items() if k == kw]
        if len(analyses) < 1:
            msg = "No analysis found for sample '${ar}' matching keyword '${kw}'"
            self.log(msg, mapping=dict(kw=kw, ar=ar.getId()))
            return None
        if len(analyses) > 1:
            self.warn(
                'Multiple analyses found matching Keyword "${kw}"',
                mapping=dict(kw=kw)
            )
            return None
        return analyses[0]

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
    def is_reference_sample(reference_sample_id):
        query = dict(portal_type="ReferenceSample", getId=reference_sample_id)
        brains = api.search(query, SENAITE_CATALOG)
        return True if brains else False

    @staticmethod
    def get_reference_sample(reference_sample_id, kw):
        query = dict(portal_type="ReferenceSample", getId=reference_sample_id)
        brains = api.search(query, SENAITE_CATALOG)
        if len(brains) < 1:
            lmsg = "No reference sample found for sample {} matching Keyword {}"
            msg = lmsg.format(reference_sample_id, kw)
            raise AnalysisNotFound(msg)
        if len(brains) > 1:
            lmsg = "Multiple objects found for sample {} matching Keyword {}"
            msg = lmsg.format(reference_sample_id, kw)
            raise MultipleAnalysesFound(msg)
        return brains[0]
