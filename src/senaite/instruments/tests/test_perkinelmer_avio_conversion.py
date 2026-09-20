# -*- coding: utf-8 -*-

from unittest import TestCase

from senaite.instruments.instruments.perkinelmer.avio.avio import AvioParser


class TestAvioConversion(TestCase):

    def test_converts_results_by_element_for_any_wavelength(self):
        parser = AvioParser.__new__(AvioParser)
        parsed = {}
        cases = [
            ("U 424.167 (mg/L)", "0.687963357571", "687.963357571000"),
            ("Fe 238.863 (mg/L)", "9.08346403553", "9083.46403553000"),
            ("Ca 315.887 (mg/L)", "1.2345", "1.2345"),
            ("Ca 317.933 (mg/L)", "2.3456", "2.3456"),
            ("Mg 279.077 (mg/L)", "3.4567", "3.4567"),
            ("Ca 999.123 (mg/L)", "4.5678", "4.5678"),
            ("Mg 999.123 (mg/L)", "5.6789", "5.6789"),
            ("Zn 999.123 (mg/L)", "1.25e-3", "1.25000"),
            ("Fe 999.123 (mg/L)", "-0.002", "-2.000"),
            ("U 999.123 (mg/L)", "0", "0"),
        ]
        for header, value, expected in cases:
            keyword = parser.normalize_keyword(header)
            analysis = type("Analysis", (), {"getKeyword": keyword})()
            parser.add_result(parsed, [analysis], keyword, value, "S1", 2)
            self.assertEqual(parsed[keyword]["Result"], expected)
        self.assertEqual(len(parsed), len(cases))

    def test_converts_interim_readings(self):
        parser = AvioParser.__new__(AvioParser)
        analysis = type("Analysis", (), {
            "getKeyword": "Iron",
            "InterimFields": [{"keyword": "Fe238863"}],
        })()
        parsed = {}
        parser.add_result(parsed, [analysis], "Fe238863", "1.25", "QC1", 2)
        self.assertEqual(parsed["Iron"]["Fe238863"], "1250.00")

    def test_preserves_nonnumeric_flags(self):
        parser = AvioParser.__new__(AvioParser)
        analysis = type("Analysis", (), {"getKeyword": "Fe238863"})()
        parsed = {}
        parser.add_result(parsed, [analysis], "Fe238863", "N/A", "S1", 2)
        self.assertEqual(parsed["Fe238863"]["Result"], "N/A")
