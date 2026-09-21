# -*- coding: utf-8 -*-

import csv
from cStringIO import StringIO
from os.path import dirname, join
from unittest import TestCase

from senaite.instruments.instruments.perkinelmer.avio import avio


class Sample(object):
    def __init__(self, sample_id, weight=None):
        self.sample_id = sample_id
        self.weight = weight

    def getId(self):
        return self.sample_id

    def getAnalyses(self, getKeyword=None, full_objects=False):
        assert getKeyword == "InitialSampleWt"
        return [] if self.weight is None else [self]

    def getResult(self):
        return self.weight

    def getReferenceAnalysesGroupID(self):
        return self.sample_id


class Response(object):
    def __init__(self):
        self.headers = {}
        self.body = ""

    def setHeader(self, name, value):
        self.headers[name] = value

    def write(self, value):
        self.body += value


class TestAvioExport(TestCase):
    def setUp(self):
        self.objects = {}
        self.layout = []
        self.original_get_object = avio.api.get_object
        self.original_search = avio.api.search
        avio.api.get_object = lambda value: self.objects.get(value, value)
        avio.api.search = self.search
        self.response = Response()
        request = type("Request", (), {"RESPONSE": self.response})()
        self.export = avio.syngistixexporter(self).Export(self, request)

    def tearDown(self):
        avio.api.get_object = self.original_get_object
        avio.api.search = self.original_search

    def getId(self):
        return "WS-001"

    def getLayoutView(self):
        return self.layout

    def search(self, query, catalog):
        self.assertEqual(query["getKeyword"], "InitialSampleWt")
        sample = self.objects[query["getReferenceAnalysesGroupID"]]
        return [] if sample.weight is None else [sample]

    def add_sample(self, position, sample_id, weight=None, kind="a"):
        self.objects[sample_id] = Sample(sample_id, weight)
        item = dict(position=position, type=kind,
                    container_uid=sample_id, analysis_uid=sample_id)
        # Multiple analyses at a position must only generate one record.
        self.layout.extend([item, item.copy()])

    def test_matches_supplied_sifx_byte_for_byte(self):
        path = join(dirname(__file__), "files", "instruments", "perkinelmer",
                    "avio", "Working Lotus Syngistix Test SIFX - Local 7 cols.sifx")
        with open(path, "rb") as fixture:
            expected = fixture.read()
        for row in csv.reader(expected.splitlines()[25:]):
            self.add_sample(int(row[1]), row[2], row[3],
                            "c" if row[2].startswith("QC") else "a")
        self.layout.reverse()
        self.export([])
        self.assertEqual(self.response.body, expected)
        self.assertEqual(self.response.headers["Content-Length"], len(expected))
        self.assertIn("WS-001.sifx",
                      self.response.headers["Content-Disposition"])

    def test_utf8_missing_weights_zero_and_csv_escaping(self):
        self.add_sample(11, u'Samplé,"A"')
        self.add_sample(14, "QC-1", "0", kind="b")
        self.add_sample(20, "DUP-1", kind="d")
        self.export([])
        body = self.response.body
        self.assertIn(u'Samplé'.encode("utf-8"), body)
        self.assertIn("NumberOfDataValues=3\r\n", body)
        rows = list(csv.reader(StringIO(body)))[25:]
        self.assertEqual(rows[0], ["Data1=1", "11",
                                  u'Samplé,"A"'.encode("utf-8"),
                                  "", "", "", ""])
        self.assertEqual(rows[1], ["Data2=2", "14", "QC-1", "0", "", "", ""])
        self.assertEqual(rows[2], ["Data3=3", "20", "DUP-1", "", "", "", ""])
        self.assertEqual(self.response.headers["Content-Type"],
                         "text/csv; charset=utf-8")

    def test_empty_worksheet(self):
        self.export([])
        self.assertTrue(self.response.body.startswith("\r\n[System Description]"))
        self.assertTrue(self.response.body.endswith("NumberOfDataValues=0\r\n"))
