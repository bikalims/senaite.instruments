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
from senaite.instruments.instruments.hach.dr3900.dr3900 import dr3900import
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest


TITLE = 'Hach DR3900'
IFACE = 'senaite.instruments.instruments' \
        '.hach.dr3900.dr3900.dr3900import'

here = abspath(dirname(__file__))
path = join(here, 'files', 'instruments', 'hach', 'dr3900')

fn_general_utf8 = join(path,'dr3900_general_utf8.csv')
fn_general_utf16 = join(path,'dr3900_general_utf16.csv') #Only file in utf16
fn_no_factor = join(path, 'dr3900_no_factor.csv')
fn_many_analyses = join(path, 'dr3900_many_analyses.csv')
fn_sameID = join(path, 'dr3900_sameID_error_test.csv')
fn_no_result = join(path, 'dr3900_no_result.csv')

service_interims = [
    dict(keyword='reading', title='Reading', hidden=False)
]

calculation_interims = [
    dict(keyword='reading', title='Reading', hidden=False),
    dict(keyword='factor', title='Factor', hidden=False)
]


class TestDR3900(BaseTestCase):

    def setUp(self):
        super(TestDR3900, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ['Member', 'LabManager'])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title='Happy Hills', ClientID='HH')

        self.contact = self.add_contact(
            self.client, Firstname='Rita', Surname='Mohale')

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title='DR3900'),
            Manufacturer=self.add_manufacturer(title='Hach'),
            Supplier=self.add_supplier(title='Instruments Inc'),
            ImportDataInterface=IFACE)

        self.calculation = self.add_calculation(
            title='Dilution', Formula='[reading] * [factor]',
            InterimFields=calculation_interims)

        self.services = [
            self.add_analysisservice(
                title='Turbidity',
                Keyword='FNU',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(title='Phyisical Properties'),
                Calculation=self.calculation,
                InterimFields=service_interims),
            self.add_analysisservice(
                title='Magnesium',
                Keyword='Mg',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(title='Organic'),
                Calculation= self.calculation,
                InterimFields=service_interims)
        ]
        self.sampletype = self.add_sampletype(
            title='Dust', RetentionPeriod=dict(days=1),
            MinimumVolume='1 kg', Prefix='DU')

    def test_general_utf8(self):
        
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        
        ar_4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        api.do_transition_for(ar_4, 'receive')
        data = open(fn_general_utf8, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_general_utf8))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity_first = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_second = ar_2.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_third = ar_3.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_fourth = ar_4.getAnalyses(full_objects=True, getKeyword='FNU')[0]

        test_results = eval(results)

        
        self.assertEqual(a_turbidity_first.getResult(), '32.9')
        self.assertEqual(a_turbidity_second.getResult(), '69.0')
        self.assertEqual(a_turbidity_third.getResult(), '114.0')
        self.assertEqual(a_turbidity_fourth.getResult(), '34.3')

    def test_general_utf16(self):
        
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        
        ar_4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        api.do_transition_for(ar_4, 'receive')
        data = open(fn_general_utf16, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_general_utf16))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity_first = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_second = ar_2.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_third = ar_3.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_fourth = ar_4.getAnalyses(full_objects=True, getKeyword='FNU')[0]

        test_results = eval(results)

        
        self.assertEqual(a_turbidity_first.getResult(), '32.9')
        self.assertEqual(a_turbidity_second.getResult(), '69.0')
        self.assertEqual(a_turbidity_third.getResult(), '114.0')
        self.assertEqual(a_turbidity_fourth.getResult(), '34.3')

    def test_single_ar_with_multiple_analysis_services(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        data = open(fn_many_analyses, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_many_analyses))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_magnesium = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)

        self.assertEqual(a_turbidity.getResult(), '32.9')
        self.assertEqual(a_magnesium.getResult(), '69.0')

    def test_same_ID_appears_more_than_once(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        
        ar_4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        api.do_transition_for(ar_4, 'receive')
        data = open(fn_sameID, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_sameID))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity_first = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_second = ar_2.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_third = ar_3.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_fourth = ar_4.getAnalyses(full_objects=True, getKeyword='FNU')[0]

        test_results = eval(results)

        #No results will be updated if one of the ID's have not been entered correctly
        self.assertEqual(a_turbidity_first.getResult(), '')
        self.assertEqual(a_turbidity_second.getResult(), '')
        self.assertEqual(a_turbidity_third.getResult(), '')
        self.assertEqual(a_turbidity_fourth.getResult(), '')
    
    def test_no_factor(self):
        
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        
        ar_4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        api.do_transition_for(ar_4, 'receive')
        data = open(fn_no_factor, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_no_factor))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity_first = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_second = ar_2.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_third = ar_3.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_fourth = ar_4.getAnalyses(full_objects=True, getKeyword='FNU')[0]

        test_results = eval(results)
        
        self.assertEqual(a_turbidity_first.getResult(), '32.9')
        self.assertEqual(a_turbidity_second.getResult(), '69.0')
        self.assertEqual(a_turbidity_third.getResult(), '38.0')
        self.assertEqual(a_turbidity_fourth.getResult(), '34.3')


    def test_no_result(self):
        
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        ar_2 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        ar_3 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])
        
        ar_4 = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        api.do_transition_for(ar_4, 'receive')
        data = open(fn_no_result, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_no_result))

        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = dr3900import.Import(self.portal, request)
        a_turbidity_first = ar.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_second = ar_2.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_third = ar_3.getAnalyses(full_objects=True, getKeyword='FNU')[0]
        a_turbidity_fourth = ar_4.getAnalyses(full_objects=True, getKeyword='FNU')[0]

        test_results = eval(results)
        
        self.assertEqual(a_turbidity_first.getResult(), '32.9')
        self.assertEqual(a_turbidity_second.getResult(), '')
        self.assertEqual(a_turbidity_third.getResult(), '114.0')
        self.assertEqual(a_turbidity_fourth.getResult(), '34.3')


