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
from senaite.instruments.instruments.agilent.flameatomic.flameatomic import flameatomicimport
from senaite.instruments.tests import TestFile
from senaite.instruments.tests.base import BaseTestCase
from zope.publisher.browser import FileUpload
from zope.publisher.browser import TestRequest


TITLE = 'Agilent Flame Atomic Absorption'
IFACE = 'senaite.instruments.instruments' \
        '.agilent.flameatomic.flameatomic.flameatomicimport'

here = abspath(dirname(__file__))
path = join(here, 'files', 'instruments', 'agilent', 'flameatomic')
fn_many_analyses = join(path, 'flameatomic_many_analyses.xlsx')
fn_sameID = join(path, 'flameatomic_sameID_error_test.xlsx')
fn_result_over = join(path,'flameatomic_OVER.xlsx' )
fn_many_analyses_csv = join(path, 'flameatomic_many_analyses_csv.csv')
fn_sameID_csv = join(path, 'flameatomic_sameID_error_test_csv.csv')
fn_result_over_csv = join(path,'flameatomic_OVER_csv.csv' )
fn_result_over_utf8 = join(path,'flameatomic_OVER_utf8.prn' )


service_interims = [
    dict(keyword='reading', title='Reading', hidden=False)
]

calculation_interims = [
    dict(keyword='reading', title='Reading', hidden=False),
    dict(keyword='factor', title='Factor', hidden=False)
]


class TestFlameAtomic(BaseTestCase):

    def setUp(self):
        super(TestFlameAtomic, self).setUp()
        setRoles(self.portal, TEST_USER_ID, ['Member', 'LabManager'])
        login(self.portal, TEST_USER_NAME)

        self.client = self.add_client(title='Happy Hills', ClientID='HH')

        self.contact = self.add_contact(
            self.client, Firstname='Rita', Surname='Mohale')

        self.instrument = self.add_instrument(
            title=TITLE,
            InstrumentType=self.add_instrumenttype(title='Flame Atomic Absorption'),
            Manufacturer=self.add_manufacturer(title='Agilent'),
            Supplier=self.add_supplier(title='Instruments Inc'),
            ImportDataInterface=IFACE)

        self.calculation = self.add_calculation(
            title='Dilution', Formula='[reading] * [factor]',
            InterimFields=calculation_interims)

        self.services = [
            self.add_analysisservice(
                title='Gold',
                Keyword='Au',
                PointOfCapture='lab',
                Category=self.add_analysiscategory(title='Organic'),
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
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_mag = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)  # noqa
        self.assertEqual(a_gold.getResult(), '0.75')
        self.assertEqual(a_mag.getResult(), '3.0')


    def test_OVER_result(self):
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

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        data = open(fn_result_over, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_result_over))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)
        self.assertEqual(a_gold_first.getResult(), '0.5')
        self.assertEqual(a_gold_second.getResult(), '0.375')
        self.assertEqual(a_gold_third.getResult(), '999999.0')


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

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        data = open(fn_sameID, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_sameID))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)
        #No results will be updated if one of the ID's have not been entered correctly
        self.assertEqual(a_gold_first.getResult(), '')
        self.assertEqual(a_gold_second.getResult(), '')
        self.assertEqual(a_gold_third.getResult(), '')

    # CSV testing

    def test_single_ar_with_multiple_analysis_services_csv(self):
        ar = self.add_analysisrequest(
            self.client,
            dict(Client=self.client.UID(),
                 Contact=self.contact.UID(),
                 DateSampled=datetime.now().date().isoformat(),
                 SampleType=self.sampletype.UID()),
            [srv.UID() for srv in self.services])

        api.do_transition_for(ar, 'receive')
        data = open(fn_many_analyses_csv, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_many_analyses_csv))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_mag = ar.getAnalyses(full_objects=True, getKeyword='Mg')[0]
        test_results = eval(results)  # noqa
        self.assertEqual(a_gold.getResult(), '0.75')
        self.assertEqual(a_mag.getResult(), '3.0')


    def test_OVER_result_csv(self):
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

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        data = open(fn_result_over_csv, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_result_over_csv))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)
        self.assertEqual(a_gold_first.getResult(), '0.5')
        self.assertEqual(a_gold_second.getResult(), '0.375')
        self.assertEqual(a_gold_third.getResult(), '999999.0')


    def test_same_ID_appears_more_than_once_csv(self):
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

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        data = open(fn_sameID_csv, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_sameID_csv))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))
            
        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)
        #No results will be updated if one of the ID's have not been entered correctly
        self.assertEqual(a_gold_first.getResult(), '')
        self.assertEqual(a_gold_second.getResult(), '')
        self.assertEqual(a_gold_third.getResult(), '')
    
    #utf8 test

    def test_OVER_result_utf8(self):
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

        api.do_transition_for(ar, 'receive')
        api.do_transition_for(ar_2, 'receive')
        api.do_transition_for(ar_3, 'receive')
        data = open(fn_result_over_utf8, 'r').read()
        import_file = FileUpload(TestFile(cStringIO.StringIO(data), fn_result_over_utf8))
        
        request = TestRequest(form=dict(
            submitted=True,
            artoapply='received_tobeverified',
            results_override='override',
            instrument_results_file=import_file,
            instrument=api.get_uid(self.instrument)))

        results = flameatomicimport.Import(self.portal, request)
        a_gold_first = ar.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_second = ar_2.getAnalyses(full_objects=True, getKeyword='Au')[0]
        a_gold_third = ar_3.getAnalyses(full_objects=True, getKeyword='Au')[0]

        test_results = eval(results)
        self.assertEqual(a_gold_first.getResult(), '0.5')
        self.assertEqual(a_gold_second.getResult(), '0.375')
        self.assertEqual(a_gold_third.getResult(), '999999.0')


def test_suite():
    suite = unittest.TestSuite()
    suite.addTest(unittest.makeSuite(TestFlameAtomic))
    return suite
