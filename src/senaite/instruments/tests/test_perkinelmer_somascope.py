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
from senaite.instruments.instruments.perkinelmer.lactoscope.lactoscopeh23061316 import (
    importer,
)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from senaite.instruments.tests.base import DataTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest

IFACE = (
    "senaite.instruments.instruments"
    ".perkinelmer.lactoscope.lactoscopeh23061316.importer"
)
TITLE = "Lactoscope H230613 16 COMP"

here = abspath(dirname(__file__))
path = join(here, "files", "instruments", "perkinelmer", "lactoscope")
fn1 = join(path, "Lactoscope_H230613_16COMP.xlsx")

service_interims = []

calculation_interims = []


class TestLactoscopeH23061316COMP(DataTestCase):
    def setUp(self):
        super(TestLactoscopeH23061316COMP, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ["Member", "LabManager"])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title="Happy Hills", ClientID="HH")

        self.contact = self.add_contact(self.client, Firstname="Rita", Surname="Mohale")

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title="5110 ICP-OES"),
            Manufacturer=self.add_manufacturer(title="ICP-OES"),
            Supplier=self.add_supplier(title="Instruments Inc"),
            ImportDataInterface=IFACE,
        )

        self.services = [
            self.add_analysisservice(
                title="Predicted Fat % m/m",
                Keyword="PredictedFat",
                PointOfCapture="lab",
                Category="Metals",
            ),
            self.add_analysisservice(
                title="Predicted Protein % m/m",
                Keyword="PredictedProtein",
                PointOfCapture="lab",
                Category="Metals",
            ),
        ]
        self.sampletype = self.add_sampletype(
            title="Dust",
            RetentionPeriod=dict(days=1),
            MinimumVolume="1 kg",
            Prefix="DU",
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
        api.do_transition_for(ar, "receive")
        # worksheet - test breaks on worksheet when adding an attachment
        # worksheet = self.add_worksheet(ar)
        # duplicate = self.add_duplicate(worksheet)
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
        pfm = ar.getAnalyses(full_objects=True, getKeyword="PredictedFat")[0]
        ppm = ar.getAnalyses(full_objects=True, getKeyword="PredictedProtein")[0]
        test_results = eval(results)  # noqa
        self.assertEqual(pfm.getResult(), "5.22")
        self.assertEqual(ppm.getResult(), "3.88")


def test_suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestLactoscopeH23061316COMP))
    return suite
