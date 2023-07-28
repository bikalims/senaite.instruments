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

import unittest2 as unittest
from plone.app.testing import TEST_USER_ID
from plone.app.testing import TEST_USER_NAME
from plone.app.testing import login
from plone.app.testing import setRoles

from bika.lims import api
from senaite.instruments.instruments.bika.software.software import (
    softwareimport)
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest


TITLE = 'Bika Software'
IFACE = 'senaite.instruments.instruments' \
        '.bika.software.software.softwareimport'

here = abspath(dirname(__file__))
path = join(here, 'files', 'instruments', 'bika', 'software')

fn_single_sample = join(path, 'software_single_sample.xls')
fn_multi_sample = join(path, 'software_multi_sample.xls')


class TestSoftware(BaseTestCase):

    def setUp(self):
        super(TestSoftware, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ['Member', 'LabManager'])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title='Happy Hills', ClientID='HH')

        self.contact = self.add_contact(
            self.client, Firstname='Rita', Surname='Mohale')

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title='Software'),
            Manufacturer=self.add_manufacturer(title='Bika'),
            Supplier=self.add_supplier(title='Bika Labs'),
            ImportDataInterface=IFACE)

        self.services = [
            self.add_analysisservice(
                title='Bromide',
                Keyword='Br',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Phyisical Properties'),
                AllowManualUncertainty='True'),
            self.add_analysisservice(
                title='Silica',
                Keyword='SiO',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Phyisical Properties'),
                AllowManualUncertainty='True'),
            self.add_analysisservice(
                title='Sulfate',
                Keyword='SO',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Phyisical Properties'),
                AllowManualUncertainty='True'),
            self.add_analysisservice(
                title='Calcium',
                Keyword='Ca',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Phyisical Properties'),
                AllowManualUncertainty='True'),
            self.add_analysisservice(
                title='Magnesium',
                Keyword='Mg',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Phyisical Properties'),
                AllowManualUncertainty='True'),
            self.add_analysisservice(
                title='Chloride',
                Keyword='Cl',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(
                    title='Organic'),
                AllowManualUncertainty='True')
        ]
        self.sampletype = self.add_sampletype(
            title='Dust', RetentionPeriod=dict(days=1),
            MinimumVolume='1 kg', Prefix='DU')

    def test_single_ar_with_multiple_analysis_services(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        data = open(fn_single_sample, 'r').read()
        import_file = FileUpload(
            TestFile(cStringIO.StringIO(data), fn_single_sample))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = softwareimport.Import(self.portal, request)
        a_bromide = ar.getAnalyses(full_objects=True, getKeyword='Br')[0]
        a_chloride = ar.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        a_silica = ar.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        a_sulfate = ar.getAnalyses(full_objects=True, getKeyword='SO')[0]
        a_calcium = ar.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        a_magnesium = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)

        self.assertEqual(a_bromide.getResult(), '-1.0')
        self.assertEqual(a_chloride.getResult(), '155.0')
        self.assertEqual(a_silica.getResult(), '1000000001.0')
        self.assertEqual(a_sulfate.getResult(), '27.1')
        self.assertEqual(a_calcium.getResult(), '34.5')
        self.assertEqual(a_magnesium.getResult(), '10.3')

    def test_single_ar_uncertainty_values(self):

        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        data = open(fn_single_sample, 'r').read()
        import_file = FileUpload(TestFile(
            cStringIO.StringIO(data), fn_single_sample))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = softwareimport.Import(self.portal, request)
        a_bromide = ar.getAnalyses(full_objects=True, getKeyword='Br')[0]
        a_chloride = ar.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        a_silica = ar.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        a_sulfate = ar.getAnalyses(full_objects=True, getKeyword='SO')[0]
        a_calcium = ar.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        a_magnesium = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)

        self.assertEqual(a_bromide.getUncertainty(), None)
        self.assertEqual(a_chloride.getUncertainty(), '25.0')
        self.assertEqual(a_silica.getUncertainty(), None)
        self.assertEqual(a_sulfate.getUncertainty(), '25.0')
        self.assertEqual(a_calcium.getUncertainty(), None)
        self.assertEqual(a_magnesium.getUncertainty(), '0.010')

    def test_multi_sample_multi_service(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar2, 'receive')
        api.do_transition_for(ar3, 'receive')
        api.do_transition_for(ar4, 'receive')

        data = open(fn_multi_sample, 'r').read()
        import_file = FileUpload(TestFile(
            cStringIO.StringIO(data), fn_multi_sample))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = softwareimport.Import(self.portal, request)

        a_bromide = ar.getAnalyses(full_objects=True, getKeyword='Br')[0]
        a_chloride = ar.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        a_silica = ar.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        a_sulfate = ar.getAnalyses(full_objects=True, getKeyword='SO')[0]
        a_calcium = ar.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        a_magnesium = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        b_bromide = ar2.getAnalyses(full_objects=True, getKeyword='Br')[0]
        b_chloride = ar2.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        b_silica = ar2.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        b_sulfate = ar2.getAnalyses(full_objects=True, getKeyword='SO')[0]
        b_calcium = ar2.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        b_magnesium = ar2.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        c_bromide = ar3.getAnalyses(full_objects=True, getKeyword='Br')[0]
        c_chloride = ar3.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        c_silica = ar3.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        c_sulfate = ar3.getAnalyses(full_objects=True, getKeyword='SO')[0]
        c_calcium = ar3.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        c_magnesium = ar3.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        d_bromide = ar4.getAnalyses(full_objects=True, getKeyword='Br')[0]
        d_chloride = ar4.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        d_silica = ar4.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        d_sulfate = ar4.getAnalyses(full_objects=True, getKeyword='SO')[0]
        d_calcium = ar4.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        d_magnesium = ar4.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        test_results = eval(results)

        self.assertEqual(a_bromide.getResult(), '')
        self.assertEqual(a_chloride.getResult(), '10.8')
        self.assertEqual(a_silica.getResult(), '2.48')
        self.assertEqual(a_sulfate.getResult(), '-1.0')
        self.assertEqual(a_calcium.getResult(), '5.89')
        self.assertEqual(a_magnesium.getResult(), '1.3')

        self.assertEqual(b_bromide.getResult(), '12.3')
        self.assertEqual(b_chloride.getResult(), '155.0')
        self.assertEqual(b_silica.getResult(), '13.5')
        self.assertEqual(b_sulfate.getResult(), '27.1')
        self.assertEqual(b_calcium.getResult(), '34.5')
        self.assertEqual(b_magnesium.getResult(), '10.3')

        self.assertEqual(c_bromide.getResult(), '')
        self.assertEqual(c_chloride.getResult(), '30.1')
        self.assertEqual(c_silica.getResult(), '4.88')
        self.assertEqual(c_sulfate.getResult(), '-1.0')
        self.assertEqual(c_calcium.getResult(), '9.05')
        self.assertEqual(c_magnesium.getResult(), '2.13')

        self.assertEqual(d_bromide.getResult(), '')
        self.assertEqual(d_chloride.getResult(), '29.0')
        self.assertEqual(d_silica.getResult(), '')
        self.assertEqual(d_sulfate.getResult(), '')
        self.assertEqual(d_calcium.getResult(), '9.16')
        self.assertEqual(d_magnesium.getResult(), '2.51')

    def test_multi_sample_multi_uncertainty(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar2, 'receive')
        api.do_transition_for(ar3, 'receive')
        api.do_transition_for(ar4, 'receive')

        data = open(fn_multi_sample, 'r').read()
        import_file = FileUpload(TestFile(
            cStringIO.StringIO(data), fn_multi_sample))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = softwareimport.Import(self.portal, request)
        a_bromide = ar.getAnalyses(full_objects=True, getKeyword='Br')[0]
        a_chloride = ar.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        a_silica = ar.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        a_sulfate = ar.getAnalyses(full_objects=True, getKeyword='SO')[0]
        a_calcium = ar.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        a_magnesium = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        b_bromide = ar2.getAnalyses(full_objects=True, getKeyword='Br')[0]
        b_chloride = ar2.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        b_silica = ar2.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        b_sulfate = ar2.getAnalyses(full_objects=True, getKeyword='SO')[0]
        b_calcium = ar2.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        b_magnesium = ar2.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        c_bromide = ar3.getAnalyses(full_objects=True, getKeyword='Br')[0]
        c_chloride = ar3.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        c_silica = ar3.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        c_sulfate = ar3.getAnalyses(full_objects=True, getKeyword='SO')[0]
        c_calcium = ar3.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        c_magnesium = ar3.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        d_bromide = ar4.getAnalyses(full_objects=True, getKeyword='Br')[0]
        d_chloride = ar4.getAnalyses(full_objects=True, getKeyword='Cl')[0]
        d_silica = ar4.getAnalyses(full_objects=True, getKeyword='SiO')[0]
        d_sulfate = ar4.getAnalyses(full_objects=True, getKeyword='SO')[0]
        d_calcium = ar4.getAnalyses(full_objects=True, getKeyword='Ca')[0]
        d_magnesium = ar4.getAnalyses(full_objects=True, getKeyword='Mg')[0]

        test_results = eval(results)

        self.assertEqual(a_bromide.getUncertainty(), None)
        self.assertEqual(a_chloride.getUncertainty(), '5.0')
        self.assertEqual(a_silica.getUncertainty(), '1.00')
        self.assertEqual(a_sulfate.getUncertainty(), None)
        self.assertEqual(a_calcium.getUncertainty(), '0.030')
        self.assertEqual(a_magnesium.getUncertainty(), '0.010')

        self.assertEqual(b_bromide.getUncertainty(), '1.0')
        self.assertEqual(b_chloride.getUncertainty(), '25.0')
        self.assertEqual(b_silica.getUncertainty(), '2.50')
        self.assertEqual(b_sulfate.getUncertainty(), '25.0')
        self.assertEqual(b_calcium.getUncertainty(), '0.040')
        self.assertEqual(b_magnesium.getUncertainty(), '0.010')

        self.assertEqual(c_bromide.getUncertainty(), None)
        self.assertEqual(c_chloride.getUncertainty(), '5.0')
        self.assertEqual(c_silica.getUncertainty(), '1.00')
        self.assertEqual(c_sulfate.getUncertainty(), None)
        self.assertEqual(c_calcium.getUncertainty(), '0.050')
        self.assertEqual(c_magnesium.getUncertainty(), '0.010')

        self.assertEqual(d_bromide.getUncertainty(), None)
        self.assertEqual(d_chloride.getUncertainty(), '5.0')
        self.assertEqual(d_silica.getUncertainty(), None)
        self.assertEqual(d_sulfate.getUncertainty(), None)
        self.assertEqual(d_calcium.getUncertainty(), '0.060')
        self.assertEqual(d_magnesium.getUncertainty(), '0.010')
