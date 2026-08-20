# -*- coding: utf-8 -*-

import cStringIO
from datetime import datetime
from os.path import abspath
from os.path import dirname
from os.path import join

from plone.app.testing import TEST_USER_ID
from plone.app.testing import TEST_USER_NAME
from plone.app.testing import login
from plone.app.testing import setRoles
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest

from bika.lims import api
from senaite.instruments.instruments.perkinelmer.avio.avio import AvioParser
from senaite.instruments.instruments.perkinelmer.avio.avio import importer
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase


TITLE = "Perkin Elmer Syngistix Uranium"
IFACE = (
    "senaite.instruments.instruments."
    "perkinelmer.avio.avio.importer"
)
here = abspath(dirname(__file__))
test_file = join(
    here, "files", "instruments", "perkinelmer", "avio",
    "Lotus Avio Environmental 22-07 XLS.xlsx")


class TestAvio(BaseTestCase):

    def setUp(self):
        super(TestAvio, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)
        self.client = self.add_client(title="Happy Hills", ClientID="HH")
        self.contact = api.create(
            self.client, "Contact", Firstname="Rita", Surname="Mohale")
        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="Avio ICP"),
            Manufacturer=self.add_manufacturer(title="Perkin Elmer"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE,
        )
        category = self.add_analysiscategory(title="ICP")
        self.services = [
            self.add_analysisservice(
                title="Uranium 424.167", Keyword="U424167",
                PointOfCapture="lab", Category=category),
            self.add_analysisservice(
                title="Iron 238.863", Keyword="Fe238863",
                PointOfCapture="lab", Category=category),
        ]
        self.sampletype = self.add_sampletype(title="Environmental")

    def make_ar(self, sample_id):
        ar = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [service.UID() for service in self.services],
        )
        ar.setId(sample_id)
        api.do_transition_for(ar, "receive")
        return ar

    def import_workbook(self):
        data = open(test_file, "rb").read()
        upload = FileUpload(TestFile(cStringIO.StringIO(data), test_file))
        request = TestRequest(form=dict(
            submitted=True,
            artoapply="received_tobeverified",
            results_override="override",
            instrument_results_file=upload,
            # Deliberately request another sheet: Avio must ignore this.
            worksheet="Corrected Intensities",
            instrument=api.get_uid(self.instrument),
        ))
        return importer.Import(self.portal, request)

    def test_header_normalization(self):
        headers = {
            "U 424.167\n(cps)": "U424167",
            "U 424.167\n(mg/L)": "U424167",
            "Fe 238.863 (cps)": "Fe238863",
            "Fe-lb 238.863 (cps)": "Fe238863",
            "Ca 315.887 (cps)": "Ca315887",
            "Ca 317.933 (cps)": "Ca317933",
            "Mg 279.077 (cps)": "Mg279077",
            "U3O8 (cps)": "U3O8",
        }
        for header, keyword in headers.items():
            self.assertEqual(
                AvioParser.normalize_keyword(header), keyword)

    def test_imports_concentration_in_sample_units(self):
        ar = self.make_ar("SW07 22.07.26")
        self.import_workbook()
        uranium = ar.getAnalyses(
            full_objects=True, getKeyword="U424167")[0]
        iron = ar.getAnalyses(
            full_objects=True, getKeyword="Fe238863")[0]
        self.assertEqual(uranium.getResult(), "0.687963357571")
        self.assertEqual(iron.getResult(), "9.08346403553")

    def test_qc_uses_reference_analysis_group_id(self):
        parser = AvioParser.__new__(AvioParser)
        searches = []
        original_search = api.search

        def search(query, catalog):
            searches.append(query)
            return [object()]

        api.search = search
        try:
            self.assertTrue(parser.is_analysis_group_id("QC10"))
        finally:
            api.search = original_search

        self.assertEqual(searches[0]["getReferenceAnalysesGroupID"], "QC10")
        self.assertEqual(
            searches[0]["portal_type"],
            ["DuplicateAnalysis", "ReferenceAnalysis"],
        )

