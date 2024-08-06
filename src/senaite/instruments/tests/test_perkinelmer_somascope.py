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
from senaite.instruments.instruments.perkinelmer.somascope.somascopeh23061316scc import (
    importer,
)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from senaite.instruments.tests.base import DataTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest

IFACE = (
    "senaite.instruments.instruments"
    ".perkinelmer.somascope.somascopeh23061316scc.importer"
)
TITLE = "Somascope H230613-16 SCC"

here = abspath(dirname(__file__))
path = join(here, "files", "instruments", "perkinelmer", "somascope")
fn1 = join(path, "Somascope_H230613-16_SCC.xlsx")

service_interims = []

calculation_interims = []


class TestSomascopeH23061316SCC(DataTestCase):
    def setUp(self):
        super(TestSomascopeH23061316SCC, self).setUp()
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
                title="DU.SCC - Somatic Cell Enumerations in Milk",
                Keyword="DU_SCC",
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
        du = ar.getAnalyses(full_objects=True, getKeyword="DU_SCC")[0]
        test_results = eval(results)  # noqa
        self.assertEqual(du.getResult(), "123.92")


def test_suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestSomascopeH23061316SCC))
    return suite
