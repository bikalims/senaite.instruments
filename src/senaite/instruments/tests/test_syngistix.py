# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS
#
# Copyright 2018 by it's authors.


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
from senaite.instruments.instruments.perkinelmer.syngistix.syngistix import (
    importer,
)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest


TITLE = "Syngistix Import Interface"
IFACE = "senaite.instruments.instruments" ".perkinelmer.syngistix.importer"

here = abspath(dirname(__file__))
path = join(here, "files", "instruments", "perkinelmer", "syngistix")

test_file = join(path, "syngistix_test_file.xlsx")

service_interims = [dict(keyword="Reading", title="Reading", hidden=False)]


class TestSyngistix(BaseTestCase):
    def setUp(self):
        super(TestSyngistix, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title="Happy Hills", ClientID="HH")

        self.contact = self.add_contact(
            self.client, Firstname="Rita", Surname="Mohale"
        )

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="Syngistix"),
            Manufacturer=self.add_manufacturer(title="Perkinelmer"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE,
        )

        self.services = [
            self.add_analysisservice(
                title="Potassium",
                Keyword="K",
                PointOfCapture="lab",
                Category=self.add_analysiscategory(title="Metals and Cations"),
                InterimFields=service_interims,
            ),
            self.add_analysisservice(
                title="Lithium",
                Keyword="Li",
                PointOfCapture="lab",
                Category=self.add_analysiscategory(title="Metals and Cations"),
                InterimFields=service_interims,
            ),
        ]
        self.sampletype = self.add_sampletype(
            title="Rock",
        )

    def test_general_utf8(self):
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
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [srv.UID() for srv in self.services],
        )

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [srv.UID() for srv in self.services],
        )

        ar_4 = self.add_analysisrequest(
            self.client,
            dict(
                Client=self.client.UID(),
                Contact=self.contact.UID(),
                DateSampled=datetime.now().date().isoformat(),
                SampleType=self.sampletype.UID(),
            ),
            [srv.UID() for srv in self.services],
        )
        ar.setId("RCK-0001")
        ar_2.setId("RCK-0002")
        ar_3.setId("RCK-0003")
        ar_4.setId("RCK-0004")

        api.do_transition_for(ar, "receive")
        api.do_transition_for(ar_2, "receive")
        api.do_transition_for(ar_3, "receive")
        api.do_transition_for(ar_4, "receive")
        data = open(test_file, "r").read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), test_file))

        request = TestRequest(
            form=dict(
                submitted=True,
                artoapply="received_tobeverified",
                results_override="override",
                instrument_results_file=import_file,
                instrument=api.get_uid(self.instrument),
            )
        )

        results = importer.Import(self.portal, request)  # noqa

        a_potassium_first = ar.getAnalyses(full_objects=True, getKeyword="K")[
            0
        ]
        a_potassium_second = ar_2.getAnalyses(
            full_objects=True, getKeyword="K"
        )[0]
        a_potassium_third = ar_3.getAnalyses(
            full_objects=True, getKeyword="K"
        )[0]
        a_potassium_fourth = ar_4.getAnalyses(
            full_objects=True, getKeyword="K"
        )[0]

        reading_value_1 = self.get_interim_result(a_potassium_first)
        reading_value_2 = self.get_interim_result(a_potassium_second)
        reading_value_3 = self.get_interim_result(a_potassium_third)
        reading_value_4 = self.get_interim_result(a_potassium_fourth)

        self.assertEqual(reading_value_1, "2.23")
        self.assertEqual(reading_value_2, "1.6")
        self.assertEqual(reading_value_3, "1.61")
        self.assertEqual(reading_value_4, "0.2")

    def get_interim_result(self, service):
        interims = service.getInterimFields()
        for interim in interims:
            if interim.get("keyword") == "Reading":
                return interim.get("value")
