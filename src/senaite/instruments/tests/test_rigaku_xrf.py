# -*- coding: utf-8 -*-

import cStringIO
import json
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
from senaite.instruments.instruments.xrf.rigaku.rigaku import importer
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase


TITLE = "Rigaku XRF"
IFACE = "senaite.instruments.instruments.xrf.rigaku.rigaku.importer"
FILE = join(abspath(dirname(__file__)), "files", "instruments", "xrf",
            "rigaku", "Rigaku XRF CSV - RIGAKU.csv")


class TestRigakuXRF(BaseTestCase):

    def setUp(self):
        super(TestRigakuXRF, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)
        self.client = self.add_client(title="Happy Hills", ClientID="HH")
        self.contact = api.create(
            self.client, "Contact", Firstname="Rita", Surname="Mohale")
        self.category = self.add_analysiscategory(title="XRF")
        self.sampletype = self.add_sampletype(title="Rock")
        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="Rigaku XRF"),
            Manufacturer=self.add_manufacturer(title="Rigaku"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE)

    def test_imports_only_column_e_analysis_from_first_block(self):
        services = [
            self.add_analysisservice(
                title="Uranium", Keyword="U3O8", PointOfCapture="lab",
                Category=self.category),
            self.add_analysisservice(
                title="Cellulose", Keyword="C6H10O5", PointOfCapture="lab",
                Category=self.category),
            self.add_analysisservice(
                title="Uranium intensity", Keyword="U-LA",
                PointOfCapture="lab", Category=self.category),
        ]
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(), Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [service.UID() for service in services])
        ar.setId("E1 TAIL 08.08 7H")
        api.do_transition_for(ar, "receive")

        response = self.import_file()

        self.assertEqual(response["errors"], [])
        analyses = dict((analysis.getKeyword(), analysis)
                        for analysis in ar.getAnalyses(full_objects=True))
        self.assertEqual(analyses["U3O8"].getResult(), "320.006")
        self.assertFalse(analyses["C6H10O5"].getResult())
        self.assertFalse(analyses["U-LA"].getResult())

    def import_file(self):
        data = open(FILE, "rb").read()
        upload = FileUpload(TestFile(cStringIO.StringIO(data), FILE))
        request = TestRequest(form=dict(
            submitted=True, artoapply="received_tobeverified",
            results_override="override", instrument_results_file=upload,
            instrument=api.get_uid(self.instrument)))
        return json.loads(importer.Import(self.portal, request))
