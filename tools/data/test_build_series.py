"""Tests for build_series.py parsing and validation. Offline, on samples in testdata/.

Run: python -m unittest tools/data/test_build_series.py   (or: python test_build_series.py)
"""
import datetime as dt
import hashlib
import json
import os
import sys
import tempfile
import io
import unittest
from contextlib import redirect_stdout
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_series as b  # noqa: E402

TD = os.path.join(HERE, "testdata")


def read(name):
    with open(os.path.join(TD, name), encoding="utf-8") as f:
        return f.read()


def le(btype, fee="1.00", first="0.0440", conv="99.90", codes=("EDO1036", "EDO1035", "EDO0936", "TOS1029")):
    return b.LeTerms(D(fee), D(first) if first else None, D(conv) if conv else None, frozenset(codes))


class MfPage(unittest.TestCase):
    def test_finds_xls_link(self):
        self.assertEqual(b.find_xls_url(read("mf_page_sample.html")),
                         "https://www.gov.pl/attachment/3d42a295-18af-43c6-b56d-de975030f64a")

    def test_missing_link_is_fatal(self):
        with self.assertRaises(b.BuildError):
            b.find_xls_url("<a href='/x'>Inny plik</a>")


class Xls(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.s = b.parse_xls(os.path.join(TD, "mf_edo_sample.xls"), types=("EDO",))

    def test_rates_margin_conversion(self):
        x = self.s["EDO1026"]
        self.assertEqual([str(r) for r in x.rates],
                         ["0.0250", "0.0330", "0.0350", "0.0440", "0.0440", "0.0700", "0.1760", "0.1160", "0.0580", "0.0440"])
        self.assertEqual(x.margin, D("0.0150"))
        self.assertEqual(x.conversion_price, D("99.70"))
        self.assertEqual(x.sale_start, dt.date(2016, 10, 1))
        self.assertEqual(x.gaps, [])

    def test_new_series_one_rate(self):
        x = self.s["EDO1036"]
        self.assertEqual(x.rates, [D("0.0535")])
        self.assertEqual(x.sale_start, dt.date(2026, 10, 1))

    def test_gap_detected(self):
        self.assertEqual(self.s["EDO1135"].gaps, [2])

    def test_missing_sheet_is_fatal(self):
        with self.assertRaises(b.BuildError):
            b.parse_xls(os.path.join(TD, "mf_edo_sample.xls"), types=("EDO", "COI"))

    def test_fixed_rate_sheet_rows(self):
        rows = [["Seria", "Kod ISIN", "Data wykupu", "Początek sprzedaży", "Koniec sprzedaży", "Cena emisyjna",
                 "Cena zamiany", "Sprzedaż", "w tym zamiana", "Oprocentowanie", "Odsetki (zł)"],
                ["TOS1029", "PL", "3 lat/a", 46296.0, 46326.0, 100.0, 99.9, "", "", 0.044, 13.79],
                ["ROS1032", "PL", "x", 46296.0, 46326.0, 100.0, "-", "", "", 0.05, ""]]   # foreign code ignored
        s = b.parse_sheet("TOS", rows, lambda v: dt.date(1899, 12, 30) + dt.timedelta(days=int(v)))
        self.assertEqual(list(s), ["TOS1029"])
        self.assertEqual(s["TOS1029"].rates, [D("0.0440")])
        self.assertIsNone(s["TOS1029"].margin)
        self.assertEqual(s["TOS1029"].sale_start, dt.date(2026, 10, 1))

    def test_dash_conversion_is_null(self):
        rows = [["Seria", "Początek sprzedaży", "Cena zamiany", "Oprocentowanie"] + [""] * 5 + ["Odsetki (zł)", "Marża"],
                ["", "", "", " w 1. roku", " w 2. roku", " w 3. roku", " w 4. roku", " w 5. roku", " w 6. roku", "", ""],
                ["ROS1032", 46296.0, "-", 0.05, "", "", "", "", "", "", 0.02]]
        s = b.parse_sheet("ROS", rows, lambda v: dt.date(1899, 12, 30) + dt.timedelta(days=int(v)))
        self.assertIsNone(s["ROS1032"].conversion_price)
        self.assertEqual(s["ROS1032"].margin, D("0.0200"))

    def test_too_precise_value_refused(self):
        with self.assertRaises(b.BuildError):
            b.dec(0.05125, 4)


class Le(unittest.TestCase):
    def test_listing_keys_by_label_and_strips_suffix(self):
        m = b.parse_le_listing(read("le_listing_sample.html"))
        self.assertEqual(m["EDO0829"], "/listy-emisyjne/?id=edo07829")   # site typo in id
        self.assertEqual(m["EDO1028"], "/listy-emisyjne/?id=edo01028")
        self.assertEqual(m["OTS0826"], "/archiwum-listow-emisyjnych/?id=ots0826")

    def test_pdf_link(self):
        self.assertEqual(b.parse_le_pdf_link(read("le_listing_sample.html")),
                         "https://www.obligacjeskarbowe.pl/media_files/4f5f6353-fdcd-4b7e-b44f-63088098c534.pdf")

    def test_ror_fee_two_clauses_same_amount(self):
        t = b.parse_le_text(read("le_ror1027.txt"), "ROR")
        self.assertEqual((t.fee, t.first_rate, t.conversion_price), (D("0.50"), D("0.0400"), D("99.90")))

    def test_coi(self):
        t = b.parse_le_text(read("le_coi1026.txt"), "COI")
        self.assertEqual((t.fee, t.first_rate, t.conversion_price), (D("0.70"), D("0.0700"), D("99.90")))

    def test_tos_fixed_wording(self):
        t = b.parse_le_text(read("le_tos1026.txt"), "TOS")
        self.assertEqual((t.fee, t.first_rate), (D("0.70"), D("0.0685")))

    def test_ots_no_fee(self):
        t = b.parse_le_text(read("le_ots0127.txt"), "OTS")
        self.assertEqual((t.fee, t.first_rate, t.conversion_price), (D("0.00"), D("0.0200"), D("100.00")))

    def test_rod_no_conversion(self):
        t = b.parse_le_text(read("le_rod1028.txt"), "ROD")
        self.assertEqual((t.fee, t.first_rate, t.conversion_price), (D("2.00"), D("0.0300"), None))

    def test_pypdf_diacritic_spaces_and_whole_zloty(self):
        t = b.parse_le_text("W pierwszym okresie odsetkowym stopa procentowa wynos 6,90% w skali roku. "
                            "jest pomniejszana o kwot ę narosłych odsetek, ale nie wy ższ ą ni ż 3 zł (trzy złote)", "ROD")
        self.assertEqual((t.fee, t.first_rate), (D("3.00"), D("0.0690")))

    def test_ambiguous_fee_is_fatal(self):
        with self.assertRaises(b.BuildError):
            b.parse_le_text("nie wyższą niż 0,70 zł ... o kwotę 0,50 zł", "COI")


class UnknownTypes(unittest.TestCase):
    TODAY = dt.date(2026, 10, 22)

    def test_new_type_in_scope_is_reported(self):
        codes = ["EDO1036", "DOS0524", "TOZ0725", "IR0995", "XYZ1029", "XYZ0826", "ABC0127"]
        self.assertEqual(b.find_unknown_types(codes, self.TODAY), {"ABC": ["ABC0127"], "XYZ": ["XYZ1029"]})

    def test_historical_types_ignored(self):
        codes = ["DOS0524", "TOZ0725", "KOS1016", "POS0323", "IR0300", "RS1099", "TZ0215", "SP1210", "Seria", "PPJ1"]
        self.assertEqual(b.find_unknown_types(codes, self.TODAY), {})

    def test_1990s_codes(self):
        self.assertEqual(b.code_maturity_month("IR0995", self.TODAY), (1995, 9))
        self.assertEqual(b.code_maturity_month("ROD1038", self.TODAY), (2038, 10))

    def test_listing_labels_with_two_letter_prefix(self):
        m = b.parse_le_listing('<option value="/listy-emisyjne/?id=nw1229,nw">NW1229</option>')
        self.assertEqual(b.find_unknown_types(m, self.TODAY), {"NW": ["NW1229"]})


class Validation(unittest.TestCase):
    def x(self, code="EDO1036", btype="EDO", start=dt.date(2026, 10, 1), rates=("0.0535",), margin="0.0200", conv="99.90"):
        return b.XlsSeries(code, btype, start, [D(r) for r in rates], D(margin) if margin else None, D(conv) if conv else None)

    def test_ok(self):
        self.assertEqual(b.validate_series(self.x(), le("EDO", "3.00", "0.0535"), dt.date(2026, 10, 9), {}), [])

    def test_stale_rates(self):
        x = self.x("EDO1035", start=dt.date(2025, 10, 1), rates=("0.0560",))
        errs = b.validate_series(x, le("EDO", "3.00", "0.0560"), dt.date(2026, 10, 9), {})
        self.assertTrue(any("expected 2" in e for e in errs), errs)

    def test_one_early_rate_allowed(self):
        x = self.x("EDO1035", start=dt.date(2025, 10, 1), rates=("0.0560", "0.0500"))
        self.assertEqual(b.validate_series(x, le("EDO", "3.00", "0.0560"), dt.date(2026, 9, 28), {}), [])

    def test_rate_out_of_range_and_missing_margin(self):
        errs = b.validate_series(self.x(rates=("0.2600",), margin=None), le("EDO", "3.00", "0.2600"), dt.date(2026, 10, 9))
        self.assertTrue(any("outside" in e for e in errs))
        self.assertTrue(any("margin missing" in e for e in errs))

    def test_le_mismatch(self):
        errs = b.validate_series(self.x(), le("EDO", "3.00", "0.0500", "99.70"), dt.date(2026, 10, 9), {})
        self.assertEqual(len(errs), 2, errs)

    def test_code_vs_sale_month(self):
        errs = b.validate_series(self.x(code="EDO0936"), le("EDO", "3.00", "0.0535"), dt.date(2026, 10, 9))
        self.assertTrue(any("code does not match" in e for e in errs))

    def test_missing_le(self):
        self.assertTrue(b.validate_series(self.x(), None, dt.date(2026, 10, 9)))

    def test_scope(self):
        today = dt.date(2026, 10, 22)
        self.assertTrue(b.in_scope("EDO1026", today))
        self.assertTrue(b.in_scope("EDO0926", today))      # 1-month grace after the last maturity month
        self.assertFalse(b.in_scope("EDO0826", today))
        self.assertTrue(b.in_scope("OTS0127", today))

    def test_expected_periods_ror(self):
        self.assertEqual(b.expected_known_periods("ROR", dt.date(2026, 8, 1), dt.date(2026, 10, 9)), 3)
        self.assertEqual(b.expected_known_periods("ROR", dt.date(2025, 10, 1), dt.date(2026, 10, 9)), 12)

    def test_tos_expanded_to_all_periods(self):
        e = b.to_json_entry(self.x("TOS1029", "TOS", rates=("0.0440",), margin=None), le("TOS"))
        self.assertEqual(e["rates"], ["0.0440"] * 3)
        self.assertIsNone(e["margin"])
        self.assertTrue(e["firstRateFixed"])

    def test_anchors(self):
        s = {"EDO1036": {"rates": ["0.0535"], "earlyRedemptionFee": "3.00", "margin": "0.0200", "conversionPrice": "99.90"}}
        errs, skipped = b.check_anchors(s, dt.date(2026, 12, 5))
        self.assertIn("EDO1026", skipped)                 # matured Oct 2026 + 1 month grace
        self.assertFalse([e for e in errs if "EDO1036" in e])
        self.assertTrue([e for e in errs if "TOS1029" in e])   # missing -> error


class VersionManifest(unittest.TestCase):
    def test_next_version(self):
        d = dt.date(2026, 11, 10)
        self.assertEqual(b.next_version(None, d), 2026111000)
        self.assertEqual(b.next_version(2026102200, d), 2026111000)
        self.assertEqual(b.next_version(2026111000, d), 2026111001)
        self.assertEqual(b.next_version(2026111299, d), 2026111300)   # previous dated ahead of the clock

    def test_same_content_ignores_volatile(self):
        a = {"schema": 1, "version": 1, "generatedAt": "x", "source": "a", "series": {"A": 1}}
        self.assertTrue(b.same_content(a, dict(a, version=2, generatedAt="y", source="b")))
        self.assertFalse(b.same_content(a, dict(a, series={"A": 2})))

    def test_manifest_hash(self):
        body = b.serialize({"schema": 1, "version": 5, "series": {}})
        m = b.build_manifest(5, "series.json", body)
        self.assertEqual(m["files"]["series.json"], {"sha256": hashlib.sha256(body).hexdigest(), "bytes": len(body)})
        self.assertEqual((m["version"], m["schema"]), (5, 1))

    def test_atomic_write(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "sub", "f.json")
            b.atomic_write(p, b"{}\n")
            with open(p, encoding="utf-8") as f:
                self.assertEqual(json.load(f), {})
            self.assertEqual(os.listdir(os.path.dirname(p)), ["f.json"])


class Index(unittest.TestCase):
    def xs(self, code, btype, start, rates, margin):
        return b.XlsSeries(code, btype, start, [D(r) for r in rates], D(margin), None)

    def test_consistent_group(self):
        idx, errs = b.build_index([self.xs("EDO1035", "EDO", dt.date(2025, 10, 1), ["0.0560", "0.0500"], "0.0200"),
                                   self.xs("COI1029", "COI", dt.date(2025, 10, 1), ["0.0535", "0.0450"], "0.0150")])
        self.assertEqual(errs, [])
        self.assertEqual(idx["CPI"]["2026-10"], D("0.0300"))

    def test_inconsistent_group_fails(self):
        _, errs = b.build_index([self.xs("EDO1035", "EDO", dt.date(2025, 10, 1), ["0.0560", "0.0500"], "0.0200"),
                                 self.xs("ROS1031", "ROS", dt.date(2025, 10, 1), ["0.0560", "0.0510"], "0.0200")])
        self.assertEqual(len(errs), 1)
        self.assertIn("CPI index 2026-10", errs[0])

    def test_groups_separate_and_period1_ignored(self):
        _, errs = b.build_index([self.xs("ROR0927", "ROR", dt.date(2026, 9, 1), ["0.0400", "0.0375"], "0.0000"),
                                 self.xs("EDO1035", "EDO", dt.date(2025, 10, 1), ["0.0999", "0.0500"], "0.0200")])
        self.assertEqual(errs, [])

    def test_validate_against_index(self):
        x = self.xs("EDO1035", "EDO", dt.date(2025, 10, 1), ["0.0560", "0.0480"], "0.0200")
        errs = b.validate_series(x, le("EDO", "3.00", "0.0560"), dt.date(2026, 10, 9), {"CPI": {"2026-10": D("0.0300")}})
        self.assertTrue(any("index" in e for e in errs), errs)


class Freshness(unittest.TestCase):
    def test_early_month_no_alarm(self):
        # 2026-11-05: XLS from 10-09 (27 days), newest XLS series 10/2026, November from LE
        e, w = b.freshness_errors(dt.date(2026, 10, 9), dt.date(2026, 10, 1), dt.date(2026, 11, 1), dt.date(2026, 11, 5))
        self.assertEqual((e, w), ([], []))

    def test_old_xls_fails(self):
        e, _ = b.freshness_errors(dt.date(2026, 10, 9), dt.date(2026, 10, 1), dt.date(2026, 11, 1), dt.date(2026, 11, 14))
        self.assertTrue(any("days old" in x for x in e))

    def test_two_months_behind_fails(self):
        e, _ = b.freshness_errors(dt.date(2026, 12, 1), dt.date(2026, 10, 1), dt.date(2026, 12, 1), dt.date(2026, 12, 2))
        self.assertTrue(any("previous month" in x for x in e))

    def test_missing_current_month(self):
        args = (dt.date(2026, 10, 30), dt.date(2026, 10, 1), dt.date(2026, 10, 1))
        e, w = b.freshness_errors(*args, dt.date(2026, 11, 2))
        self.assertEqual((e, len(w)), ([], 1))
        e, w = b.freshness_errors(*args, dt.date(2026, 11, 3))
        self.assertEqual(len(e), 1)


class LeOnly(unittest.TestCase):
    def test_from_le_text(self):
        t = b.parse_le_text(read("le_ror1027.txt"), "ROR")
        self.assertEqual(t.sale_month, dt.date(2026, 10, 1))
        self.assertEqual(t.margin, D("0.0000"))
        self.assertIn("ROR1027", t.codes)
        x = b.le_only_series("ROR1027", "ROR", t)
        self.assertEqual((x.source, x.rates, x.margin, x.conversion_price), ("LE", [D("0.0400")], D("0.0000"), D("99.90")))
        self.assertEqual(b.validate_series(x, t, dt.date(2026, 10, 9), {}, {("ROR", "2026-09"): D("0.0400")}), [])
        self.assertEqual(b.to_json_entry(x, t)["source"], "LE")

    def test_series_code(self):
        self.assertEqual(b.series_code("EDO", dt.date(2026, 11, 1)), "EDO1136")
        self.assertEqual(b.series_code("OTS", dt.date(2026, 11, 1)), "OTS0227")

    def test_wrong_le_document_rejected(self):
        t = b.parse_le_text(read("le_ror1027.txt"), "ROR")
        x = b.XlsSeries("ROR1127", "ROR", dt.date(2026, 11, 1), [D("0.0400")], D("0.0000"), D("99.90"))
        errs = b.validate_series(x, t, dt.date(2026, 11, 9), {})
        self.assertTrue(any("does not mention" in e for e in errs))
        self.assertTrue(any("sale month" in e for e in errs))


class InterestTable(unittest.TestCase):
    def test_parse_to_text(self):
        self.assertEqual(b.parse_to_text(read("to_edo1027_p10.txt")), (D("0.0490"), "EDO1027", 10))

    def test_parse_to_listing(self):
        h = ('<option value="211751,edo" data-id="edo"> EDO1027 </option>'
             '<option value="23845,211751" data-id="211751"> 2026-10-01 -&gt; 2027-10-01 </option>')
        issues, tables = b.parse_to_listing(h)
        self.assertEqual(issues, {"EDO1027": "211751"})
        self.assertEqual(tables, {"211751": {dt.date(2026, 10, 1): "23845"}})

    def test_candidate_order(self):
        mk = lambda c, start, n: b.XlsSeries(c, "EDO", start, [D("0.05")] * n, D("0.02"), None)  # noqa: E731
        scope = {c: mk(c, s, n) for c, s, n in [("EDO1027", dt.date(2017, 10, 1), 10), ("EDO1035", dt.date(2025, 10, 1), 2),
                                                 ("EDO1031", dt.date(2021, 10, 1), 6), ("EDO0936", dt.date(2026, 9, 1), 1)]}
        got = b.select_to_anchors(scope, dt.date(2026, 10, 9))["EDO"]
        self.assertEqual(got[:3], [("EDO1035", 1), ("EDO1027", 9), ("EDO1031", 5)])
        self.assertEqual(got[-1], ("EDO0936", 0))      # period 1 (fixed rate) only as last resort


class CacheSafety(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.TemporaryDirectory()
        self.old = b.CACHE
        b.CACHE = self.d.name

    def tearDown(self):
        b.CACHE = self.old
        self.d.cleanup()

    def test_signatures(self):
        self.assertTrue(b._valid(b"%PDF-1.4 ...", "pdf"))
        self.assertFalse(b._valid(b"<html>blad</html>", "pdf"))
        self.assertTrue(b._valid(bytes.fromhex("D0CF11E0A1B11AE1") + b"x", "xls"))
        self.assertFalse(b._valid(b"", None))

    def test_invalid_cached_pdf_offline_is_error(self):
        os.makedirs(os.path.join(self.d.name, "le"))
        with open(os.path.join(self.d.name, "le", "X.pdf"), "wb") as f:
            f.write(b"<html>serwis niedostepny</html>")
        with self.assertRaises(b.BuildError):
            b.fetch("http://example.invalid/x.pdf", "le/X.pdf", None, True, kind="pdf")

    def test_unreadable_pdf_leaves_no_txt(self):
        with open(os.path.join(self.d.name, "bad.pdf"), "wb") as f:
            f.write(b"%PDF-1.4 garbage")
        with self.assertRaises(b.BuildError):
            b.cached_pdf_text("bad.pdf")
        self.assertFalse(os.path.exists(os.path.join(self.d.name, "bad.txt")))


class Audit2Rules(unittest.TestCase):
    def test_html_kinds_need_content(self):
        maint = b"<html><body>Przerwa techniczna</body></html>"
        for kind in ("mf_page", "le_listing", "to_listing", "page_with_pdf"):
            self.assertFalse(b._valid(maint, kind), kind)
        self.assertTrue(b._valid(b'<html><a href="/media_files/ab12-cd.pdf">', "page_with_pdf"))
        self.assertTrue(b._valid(b'<html><option value="/listy-emisyjne/?id=edo1036,edo">', "le_listing"))
        self.assertTrue(b._valid(b'<html><option value="23845,211751">', "to_listing"))

    def test_le_only_jump_vs_previous_offer(self):
        t = b.parse_le_text(read("le_ror1027.txt"), "ROR")
        x = b.le_only_series("ROR1027", "ROR", t)
        offers = {("ROR", "2026-09"): D("0.0400")}
        self.assertEqual(b.validate_series(x, t, dt.date(2026, 10, 9), {}, offers), [])
        offers = {("ROR", "2026-09"): D("0.0600")}
        self.assertTrue(b.validate_series(x, t, dt.date(2026, 10, 9), {}, offers))
        self.assertEqual(b.validate_series(x, t, dt.date(2026, 10, 9), {}, {}), [])   # no reference: build alerts

    def test_write_alerts(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "alerts.out")
            summary = os.path.join(d, "summary.md")
            old = os.environ.get("GITHUB_STEP_SUMMARY")
            os.environ["GITHUB_STEP_SUMMARY"] = summary
            try:
                with redirect_stdout(io.StringIO()) as out:
                    b.write_alerts(["XLS stale: x", "CPI index 2026-10: withheld"], p)
                    b.write_alerts([], os.path.join(d, "empty.out"))
            finally:
                if old is None:
                    del os.environ["GITHUB_STEP_SUMMARY"]
                else:
                    os.environ["GITHUB_STEP_SUMMARY"] = old
            self.assertEqual(open(p, encoding="utf-8").read().splitlines(), ["XLS stale: x", "CPI index 2026-10: withheld"])
            self.assertEqual(os.path.getsize(os.path.join(d, "empty.out")), 0)
            self.assertIn("::warning::XLS stale: x", out.getvalue())
            self.assertIn("- CPI index 2026-10: withheld", open(summary, encoding="utf-8").read())

    def test_static_anchors_have_no_minimum(self):
        errs, skipped = b.check_anchors({}, dt.date(2040, 1, 1))
        self.assertEqual((errs, len(skipped)), ([], len(b.ANCHORS)))

    def test_withheld_anchor_skipped(self):
        errs, skipped = b.check_anchors({}, dt.date(2026, 10, 22), frozenset(b.ANCHORS))
        self.assertEqual(errs, [])


class TableSource(unittest.TestCase):
    def test_rate_without_decimals(self):
        r = b.parse_to_text("TABELA Oprocentowanie w bieżącym okresie: 4% Okres odsetkowy: 3 DOR0727 Warszawa")
        self.assertEqual(r, (D("0.0400"), "DOR0727", 3))

    def _run(self, tables):
        real = b.to_table

        def fake(code, k, start, to_listing, offline):
            key = (code, k)
            if key not in tables:
                raise b.NoTable("none")
            v = tables[key]
            if isinstance(v, Exception):
                raise v
            return D(v), "1"
        b.to_table = fake
        try:
            x = b.XlsSeries("EDO1135", "EDO", dt.date(2025, 11, 1), [D("0.0560")], D("0.0200"), D("99.90"))
            warn, added = b.augment_from_tables({"EDO1135": x}, ({}, {}), dt.date(2026, 10, 22), True)
        finally:
            b.to_table = real
        return x, warn, added

    def test_adds_next_period_and_marks_source(self):
        x, warn, added = self._run({("EDO1135", 1): "0.0610"})
        self.assertEqual((x.rates, x.to_rates, added), ([D("0.0560"), D("0.0610")], 1, ["EDO1135#2"]))
        e = b.to_json_entry(x, b.LeTerms(D("3.00"), D("0.0560"), D("99.90"), frozenset({"EDO1135"})))
        self.assertEqual(e["rateSources"], ["XLS", "TO"])
        # the XLS count check ignores table rates (XLS as of 2026-10-09 knows 1 period)
        self.assertEqual(b.validate_series(x, b.LeTerms(D("3.00"), D("0.0560"), D("99.90"), frozenset({"EDO1135"})),
                                           dt.date(2026, 10, 9), {}, {}), [])

    def test_mislabeled_table_not_used(self):
        x, warn, added = self._run({("EDO1135", 1): b.Mislabeled("table 1 is EDO1134 period 3 (mislabeled)")})
        self.assertEqual((len(x.rates), added), (1, []))
        self.assertTrue(warn)

    def test_beyond_horizon_not_fetched(self):
        x = b.XlsSeries("EDO1236", "EDO", dt.date(2026, 12, 1), [D("0.0535")], D("0.0200"), None)
        calls = []
        real = b.to_table
        b.to_table = lambda *a: calls.append(a) or (D("0.05"), "1")
        try:
            b.augment_from_tables({"EDO1236": x}, ({}, {}), dt.date(2026, 10, 22), True)
        finally:
            b.to_table = real
        self.assertEqual(calls, [])       # period 2 starts 2027-12: far beyond the look-ahead


if __name__ == "__main__":
    unittest.main()
