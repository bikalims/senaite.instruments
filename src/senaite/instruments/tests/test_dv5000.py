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

from bika.lims import api
from senaite.instruments.instruments.pg.dv5000icp.dv5000 import (
    importer,
)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest


TITLE = "PG DV5000 ICP"
IFACE = (
    "senaite.instruments.instruments."
    "pg.dv5000icp.dv5000.importer"
)

here = abspath(dirname(__file__))
path = join(
    here,
    "files",
    "instruments",
    "pg",
)

test_file = join(path, "PGDV5000ICP.xlsx")

service_interims = [
    dict(keyword="Reading", title="Reading", hidden=False),
]


class TestDV5000(BaseTestCase):

    def setUp(self):
        super(TestDV5000, self).setUp()

        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(
            title="Happy Hills",
            ClientID="HH",
        )

        self.contact = self.add_contact(
            self.client, firstname="Rita", surname="Mohale",
            # EmaailAddress="rita@lab.test",
        )

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="DV5000"),
            Manufacturer=self.add_manufacturer(title="Perkin Elmer"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE,
        )

        category = self.add_analysiscategory(
            title="ICP"
        )

        self.services = [
            self.add_analysisservice(
                title="Calcium",
                Keyword="Ca",
                PointOfCapture="lab",
                Category=category,
                InterimFields=service_interims,
            ),
            self.add_analysisservice(
                title="Magnesium",
                Keyword="Mg",
                PointOfCapture="lab",
                Category=category,
                InterimFields=service_interims,
            ),
            self.add_analysisservice(
                title="Iron",
                Keyword="Fe",
                PointOfCapture="lab",
                Category=category,
                InterimFields=service_interims,
            ),
        ]

        self.sampletype = self.add_sampletype(
            title="Rock",
        )

    def test_import(self):

        ar = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [srv.UID() for srv in self.services],
        )

        #
        # IMPORTANT
        # Your workbook contains:
        #
        #     BATCH 61 B 5/22/2026 8:23:24
        #
        # so the importer should strip the date and find
        # this AR.
        #
        ar.setId("BATCH 61 B")

        api.do_transition_for(ar, "receive")

        data = open(test_file, "rb").read()

        import_file = FileUpload(
            TestFile(
                cStringIO.StringIO(data),
                test_file,
            )
        )

        request = TestRequest(
            form=dict(
                submitted=True,
                artoapply="received_tobeverified",
                results_override="override",
                instrument_results_file=import_file,
                instrument=api.get_uid(self.instrument),
            )
        )

        importer.Import(self.portal, request)

        ca = ar.getAnalyses(
            full_objects=True,
            getKeyword="Ca",
        )[0]

        mg = ar.getAnalyses(
            full_objects=True,
            getKeyword="Mg",
        )[0]

        fe = ar.getAnalyses(
            full_objects=True,
            getKeyword="Fe",
        )[0]

        #
        # Values from the χ row
        #
        self.assertEqual(
            self.get_interim_result(ca),
            "12.34",
        )

        self.assertEqual(
            self.get_interim_result(mg),
            "4.56",
        )

        self.assertEqual(
            self.get_interim_result(fe),
            "0.78",
        )

    def get_interim_result(self, analysis):

        for interim in analysis.getInterimFields():
            if interim.get("keyword") == "Reading":
                return interim.get("value")
