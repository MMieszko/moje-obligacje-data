"""Audit tests (2026-10-22, independent reviewer). Each test encodes a property the pipeline SHOULD have;
the ones that fail document a finding. Run: python -m unittest tools/data/test_audit.py -v

Tests touching the real download cache copy what they need into a temp dir; the cache is never modified.
"""
import datetime as dt
import hashlib
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from decimal import Decimal

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_series as b  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
REAL_CACHE = os.path.join(HERE, "cache")
REPO_SERIES = os.path.join(os.path.dirname(os.path.dirname(HERE)), "data", "series.json")
HAVE_CACHE = os.path.exists(os.path.join(REAL_CACHE, "mf.xls")) and os.path.exists(REPO_SERIES)


def cache_today(days_after_xls: int = 13) -> str:
    """Clock-independent --today: a fixed distance after the cached XLS date (2026-10-09 + 13 = 2026-10-22)."""
    from email.utils import parsedate_to_datetime
    meta = json.load(open(os.path.join(REAL_CACHE, "mf.xls.meta.json"), encoding="utf-8"))
    return (parsedate_to_datetime(meta["lastModified"]).date() + dt.timedelta(days=days_after_xls)).isoformat()


def run_main(argv):
    out, err = io.StringIO(), io.StringIO()
    with redirect_stdout(out), redirect_stderr(err):
        try:
            rc = b.main(argv)
        except Exception as e:  # noqa: BLE001
            rc = ("RAISED", type(e).__name__, str(e))
    return rc, out.getvalue(), err.getvalue()


class A1_StaleXlsPublishedAsFresh(unittest.TestCase):
    """XLS two monthly updates old (Last-Modified + 62 days): the build must not publish it silently.
    Since audit 2 staleness is failure class (B): data stays published (it is old, not wrong) and an ALERT is
    raised (alerts.out -> failed 'alerts' job -> e-mail). Clock-independent: --today is derived from the
    cached XLS date, not from the wall clock (audit 2, W2)."""

    @unittest.skipUnless(HAVE_CACHE, "needs real cache")
    def test_two_month_old_xls_must_alert(self):
        meta = json.load(open(os.path.join(REAL_CACHE, "mf.xls.meta.json"), encoding="utf-8"))
        from email.utils import parsedate_to_datetime
        today = parsedate_to_datetime(meta["lastModified"]).date() + dt.timedelta(days=62)
        with tempfile.TemporaryDirectory() as d:
            shutil.copy(REPO_SERIES, os.path.join(d, "series.json"))
            alerts = os.path.join(d, "alerts.out")
            rc, out, err = run_main(["--offline", "--today", today.isoformat(), "--out", os.path.join(d, "series.json"),
                                     "--alerts-out", alerts])
            text = open(alerts, encoding="utf-8").read() if os.path.exists(alerts) else ""
            self.assertIn("XLS stale", text, f"stale XLS on {today} raised no alert; rc={rc}; stdout tail: {out[-300:]}")


class A2_NoPlausibilityForLaterPeriods(unittest.TestCase):
    """A wrong rate in period >= 2 (in range, right count, LE period-1 matches) passes every check.
    Rate - margin for the same start month is identical across all CPI-linked series (COI/EDO/ROS/ROD)
    and across ROR/DOR (NBP) - verified on the current data: 0 inconsistencies in 133 months."""

    def test_wrong_period5_rate_detected(self):
        good = b.XlsSeries("EDO0130", "EDO", dt.date(2020, 1, 1),
                           [Decimal(x) for x in "0.0270 0.0170 0.0590 0.1900 0.0840 0.0670 0.0660".split()],
                           Decimal("0.0100"), Decimal("99.90"))
        bad = b.XlsSeries(**{**good.__dict__, "rates": good.rates[:4] + [Decimal("0.0480")] + good.rates[5:]})
        le = b.LeTerms(Decimal("2.00"), Decimal("0.0270"), Decimal("99.90"))
        errs = b.validate_series(bad, le, dt.date(2026, 10, 9))
        self.assertTrue(errs, "EDO0130 with a wrong 5th-period rate passed validate_series")


class A3_AnchorsExpireWithoutReplacement(unittest.TestCase):
    def test_full_rate_anchor_still_active_after_expiry(self):
        for day in (dt.date(2026, 11, 5), dt.date(2027, 2, 5)):
            active = [c for c, a in b.ANCHORS.items() if "rates" in a and b.in_scope(c, day)]
            self.assertTrue(active, f"{day}: no multi-period anchor left in scope; only rate1/fee anchors remain")

    def test_some_anchor_per_type(self):
        day = dt.date(2027, 2, 5)
        types = {c[:3] for c in b.ANCHORS if b.in_scope(c, day)}
        self.assertEqual(set(b.TYPES), types, f"types without any anchor on {day}: {set(b.TYPES) - types}")


class A4_CrlfCheckoutBreaksManifestHash(unittest.TestCase):
    """core.autocrlf=true and no .gitattributes: a Windows checkout of series.json has CRLF. The UNCHANGED path
    hashes the on-disk CRLF bytes, git commits/publishes LF -> sha256 in manifest != served file."""

    @unittest.skipUnless(HAVE_CACHE, "needs real cache")
    def test_manifest_hash_is_of_canonical_bytes(self):
        with tempfile.TemporaryDirectory() as d:
            lf = open(REPO_SERIES, "rb").read().replace(b"\r\n", b"\n")
            with open(os.path.join(d, "series.json"), "wb") as f:
                f.write(lf.replace(b"\n", b"\r\n"))
            rc, out, err = run_main(["--offline", "--today", cache_today(), "--out", os.path.join(d, "series.json"),
                                     "--alerts-out", os.path.join(d, "alerts.out")])
            self.assertEqual(rc, 0, err[-500:])
            man = json.load(open(os.path.join(d, "manifest.json")))
            self.assertEqual(man["files"]["series.json"]["sha256"], hashlib.sha256(lf).hexdigest(),
                             "manifest hashes CRLF bytes; the LF file git publishes will fail the phone's sha check")


class A5_PdfFailurePoisonsCache(unittest.TestCase):
    """HTTP 200 with an HTML body cached as le/<code>.pdf: pypdf raises a non-BuildError (traceback, not the
    FATAL/ERROR path), and the txt file is opened before pdf_text() runs -> an EMPTY .txt stays in the cache
    forever (max_age None), so every later run fails until someone deletes it by hand."""

    @unittest.skipUnless(HAVE_CACHE, "needs real cache")
    def test_bad_pdf(self):
        code = "EDO1036"
        with tempfile.TemporaryDirectory() as d:
            cache = os.path.join(d, "cache")
            os.makedirs(os.path.join(cache, "le"))
            for n in ("mf_page.html", "mf.xls", "le_listing_0.html", "le_listing_1.html"):
                for suf in ("", ".meta.json"):
                    shutil.copy(os.path.join(REAL_CACHE, n + suf), os.path.join(cache, n + suf))
            shutil.copy(os.path.join(REAL_CACHE, "le", code + ".html"), os.path.join(cache, "le", code + ".html"))
            with open(os.path.join(cache, "le", code + ".pdf"), "wb") as f:
                f.write(b"<html><body>Przepraszamy, serwis chwilowo niedostepny</body></html>")
            real_parse = b.parse_xls
            old_cache = b.CACHE
            b.CACHE = cache
            b.parse_xls = lambda p, types=tuple(b.TYPES): {k: v for k, v in real_parse(p).items() if k == code}
            try:
                rc, out, err = run_main(["--offline", "--out", os.path.join(d, "series.json"), "--alerts-out", os.path.join(d, "alerts.out")])
            finally:
                b.CACHE, b.parse_xls = old_cache, real_parse
            self.assertEqual(rc, 1, f"expected a reported validation error, got {rc}")
            txt = os.path.join(cache, "le", code + ".txt")
            self.assertFalse(os.path.exists(txt) and os.path.getsize(txt) == 0, "empty .txt left in cache")


class A6_LeTextNotBoundToSeries(unittest.TestCase):
    """The fee is taken from whatever PDF the LE page links first; nothing checks that the PDF is the list
    for that series. Cross-checks (rate 1, conversion price) are shared by neighbouring series."""

    def test_le_of_other_series_rejected(self):
        txt = open(os.path.join(HERE, "testdata", "le_tos1026.txt"), encoding="utf-8").read()
        x = b.XlsSeries("TOS1126", "TOS", dt.date(2023, 11, 1), [Decimal("0.0685")], None, Decimal("99.90"))
        le = b.parse_le_text(txt, "TOS")
        self.assertTrue(b.validate_series(x, le, dt.date(2026, 10, 9)),
                        "LE of TOS1026 accepted for TOS1126 (fee would come from the wrong list)")


class A7_VersionFromFutureIsFatalForever(unittest.TestCase):
    def test_prev_version_ahead_of_clock(self):
        self.assertGreater(b.next_version(2026102500, dt.date(2026, 10, 22)), 2026102500)


if __name__ == "__main__":
    unittest.main()
