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

    def test_u3o8_header_matches_keyword_containing_u3o8(self):
        service = self.add_analysisservice(
            title="Uranium solution", Keyword="U3O8_sol",
            PointOfCapture="lab", Category=self.category)
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(), Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [service.UID()])
        ar.setId("AMIS0087")
        api.do_transition_for(ar, "receive")

        response = self.import_file()

        self.assertEqual(response["errors"], [])
        analysis = ar.getAnalyses(
            full_objects=True, getKeyword="U3O8_sol")[0]
        self.assertEqual(analysis.getResult(), "242.431")

    def test_u3o8_header_does_not_import_when_multiple_analyses_match(self):
        services = [
            self.add_analysisservice(
                title="Uranium solids", Keyword="U3O8_solids",
                PointOfCapture="lab", Category=self.category),
            self.add_analysisservice(
                title="Uranium slurry", Keyword="U3O8_slurry",
                PointOfCapture="lab", Category=self.category),
        ]
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(), Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [service.UID() for service in services])
        ar.setId("AMIS0087")
        api.do_transition_for(ar, "receive")

        response = self.import_file()

        analyses = ar.getAnalyses(full_objects=True)
        self.assertTrue(all(not analysis.getResult() for analysis in analyses))
        self.assertTrue(any(
            "Duplicate U3O8 Analyses found, please capture manually" in warning
            for warning in response["warns"]))

    def import_file(self):
        data = open(FILE, "rb").read()
        upload = FileUpload(TestFile(cStringIO.StringIO(data), FILE))
        request = TestRequest(form=dict(
            submitted=True, artoapply="received_tobeverified",
            results_override="override", instrument_results_file=upload,
            instrument=api.get_uid(self.instrument)))
        return json.loads(importer.Import(self.portal, request))
