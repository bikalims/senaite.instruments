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
from senaite.instruments.instruments.xrf.axios.axios import importer
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase


TITLE = "Malvern Panalytical Axios XRF"
IFACE = "senaite.instruments.instruments.xrf.axios.axios.importer"
FILE = join(abspath(dirname(__file__)), "files", "instruments", "xrf",
            "axios", "axios.xlsx")


class TestAxiosXRF(BaseTestCase):

    def setUp(self):
        super(TestAxiosXRF, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)
        self.client = self.add_client(title="Happy Hills", ClientID="HH")
        # BaseTestCase.add_contact uses the legacy edit API, which does not
        # accept the Dexterity Contact's Surname field.
        self.contact = api.create(
            self.client, "Contact", Firstname="Rita", Surname="Mohale")
        self.category = self.add_analysiscategory(title="XRF")
        self.sampletype = self.add_sampletype(title="Rock")
        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="Axios XRF"),
            Manufacturer=self.add_manufacturer(title="Malvern Panalytical"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE)

    def test_result_header_matches_analysis_keyword(self):
        service = self.add_analysisservice(
            title="Uranium", Keyword="U3O8", PointOfCapture="lab",
            Category=self.category,
            InterimFields=[
                dict(keyword="SumOfConc", title="Sum of concentration",
                     hidden=False),
            ])
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(), Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [service.UID()])
        ar.setId("AMIS0087")
        api.do_transition_for(ar, "receive")

        data = open(FILE, "rb").read()
        upload = FileUpload(TestFile(cStringIO.StringIO(data), FILE))
        request = TestRequest(form=dict(
            submitted=True, artoapply="received_tobeverified",
            results_override="override", instrument_results_file=upload,
            instrument=api.get_uid(self.instrument)))
        importer.Import(self.portal, request)

        analysis = ar.getAnalyses(full_objects=True, getKeyword="U3O8")[0]
        self.assertEqual(analysis.getResult(), "242.431")
        interims = dict((item["keyword"], item.get("value"))
                        for item in analysis.getInterimFields())
        self.assertEqual(interims["SumOfConc"], "86.186")
