# -*- coding: utf-8 -*-
#
# This file is part of SENAITE.INSTRUMENTS.
#
# SENAITE.INSTRUMENTS is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by the Free
# Software Foundation, version 2.
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
# Copyright 2018-2021 by it's authors.
# Some rights reserved, see README and LICENSE.

import cStringIO
from datetime import datetime
from os.path import abspath
from os.path import dirname
from os.path import join

import unittest2 as unittest
from plone.app.testing import TEST_USER_ID
from plone.app.testing import TEST_USER_NAME
from plone.app.testing import login
from plone.app.testing import setRoles

from bika.lims import api
from senaite.instruments.instruments.xrf.arlperformx4200xrf.alsxrf import (
    importer,
)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from senaite.instruments.tests.base import DataTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest

IFACE = (
    "senaite.instruments.instruments"
    ".xrf.arlperformx4200xrf.alsxrf.importer"
)
TITLE = "ALS ThermoFisher ARL Perform’X 4200 XRF"

here = abspath(dirname(__file__))
path = join(here, "files", "instruments", "xrf", "arlperformx4200xrf")
fn1 = join(path, "Results XRF.xlsx")

service_interims = [dict(keyword="Reading", title="Reading", hidden=False)]

calculation_interims = []


class TestALSXRF(DataTestCase):
    def setUp(self):
        super(TestALSXRF, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title="Happy Hills", ClientID="HH")

        self.contact = self.add_contact(self.client, Firstname="Rita",
                                        Surname="Mohale")

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="ALS XRF"),
            Manufacturer=self.add_manufacturer(title="ThermoFisher"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE,
        )

        self.services = [
            self.add_analysisservice(
                title="Tantalum",
                Keyword="Ta",
                PointOfCapture="lab",
                Category=self.add_analysiscategory(title="Metals and Cations"),
                InterimFields=service_interims,
            ),
            self.add_analysisservice(
                title="Tin",
                Keyword="Sn",
                PointOfCapture="lab",
                Category=self.add_analysiscategory(title="Metals and Cations"),
                InterimFields=service_interims,
            ),
        ]

        self.sampletype = self.add_sampletype(
            title="Rock",
        )

    def test_import_xlsx(self):
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
        data = open(fn1, "rb").read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn1))
        request = TestRequest(
            form=dict(
                instrument_results_file_format="xlsx",
                submitted=True,
                artoapply="received_tobeverified",
                results_override="override",
                instrument_results_file=import_file,
                instrument="",
            )
        )
        results = importer.Import(self.portal, request)
        a_tantalam_first = ar.getAnalyses(full_objects=True, getKeyword="Ta")[
            0
        ]
        a__tantalam_second = ar_2.getAnalyses(
            full_objects=True, getKeyword="Ta"
        )[0]
        a__tantalam_third = ar_3.getAnalyses(
            full_objects=True, getKeyword="Ta"
        )[0]
        a__tantalam_fourth = ar_4.getAnalyses(
            full_objects=True, getKeyword="Ta"
        )[0]

        reading_value_1 = self.get_interim_result(a_tantalam_first)
        reading_value_2 = self.get_interim_result(a__tantalam_second)
        reading_value_3 = self.get_interim_result(a__tantalam_third)
        reading_value_4 = self.get_interim_result(a__tantalam_fourth)
        test_results = eval(results)
        self.assertEqual(reading_value_1, "0.0041")
        self.assertEqual(reading_value_2, "0.0033")
        self.assertEqual(reading_value_3, "25.0")
        self.assertEqual(reading_value_4, "1.123")

    def get_interim_result(self, service):
        interims = service.getInterimFields()
        for interim in interims:
            if interim.get("keyword") == "Reading":
                return interim.get("value")
