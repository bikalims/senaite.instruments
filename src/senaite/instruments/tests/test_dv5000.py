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

test_file = join(path, "div5000icp", "PGDV5000ICP.xlsx")


class TestDV5000(BaseTestCase):

    def setUp(self):
        super(TestDV5000, self).setUp()

        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(
            title="Happy Hills",
            ClientID="HH",
        )

        self.contact = api.create(
            self.client, "Contact", Firstname="Rita", Surname="Mohale")

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
                Keyword="Ca396847-A",
                PointOfCapture="lab",
                Category=category,
            ),
            self.add_analysisservice(
                title="Magnesium",
                Keyword="Mg280271-A",
                PointOfCapture="lab",
                Category=category,
            ),
            self.add_analysisservice(
                title="Iron",
                Keyword="Fe259940-A",
                PointOfCapture="lab",
                Category=category,
            ),
            self.add_analysisservice(
                title="Calculated calcium",
                Keyword="CalculatedCa",
                PointOfCapture="lab",
                Category=category,
                InterimFields=[
                    dict(keyword="Ca396847-A", title="Calcium reading",
                         hidden=False),
                ],
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
        #     BATCH 61 B 5/22/2026 8:25:23
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
            getKeyword="Ca396847-A",
        )[0]

        mg = ar.getAnalyses(
            full_objects=True,
            getKeyword="Mg280271-A",
        )[0]

        fe = ar.getAnalyses(
            full_objects=True,
            getKeyword="Fe259940-A",
        )[0]

        #
        # Values from the χ row
        #
        self.assertEqual(ca.getResult(), "-0.067")

        self.assertEqual(mg.getResult(), "1.526")

        self.assertEqual(fe.getResult(), "1.323")

    def test_import_interim_used_by_calculation(self):
        interim_keyword = "K769896-A"
        calculation = api.create(
            self.portal.setup.calculations, "Calculation",
            title="Potassium calculation",
            Formula="[{}] * 2".format(interim_keyword),
            InterimFields=[
                dict(keyword=interim_keyword, title="Potassium reading",
                     hidden=False),
            ],
        )
        service = self.add_analysisservice(
            title="Calculated potassium",
            Keyword="CalculatedK",
            PointOfCapture="lab",
            Category=self.services[0].getCategory(),
            Calculation=calculation,
            InterimFields=[
                dict(keyword=interim_keyword, title="Potassium reading",
                     hidden=False),
            ],
        )
        ar = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [service.UID()],
        )
        ar.setId("BATCH 61 B")
        api.do_transition_for(ar, "receive")

        data = open(test_file, "rb").read()
        import_file = FileUpload(
            TestFile(cStringIO.StringIO(data), test_file))
        request = TestRequest(form=dict(
            submitted=True,
            artoapply="received_tobeverified",
            results_override="override",
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument),
        ))

        importer.Import(self.portal, request)

        analysis = ar.getAnalyses(
            full_objects=True, getKeyword="CalculatedK")[0]
        interims = dict((field["keyword"], field.get("value"))
                        for field in analysis.getInterimFields())
        self.assertEqual(interims[interim_keyword], "0.084")
        self.assertEqual(analysis.getResult(), "0.168")
