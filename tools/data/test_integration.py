"""Integration tests: full --offline builds on the real download cache (skipped without it). Local only - CI runs
test_build_series.py before the build and must not be gated by an offline build (audit 2, W2).

Run: python -m unittest tools/data/test_integration.py
"""
import datetime as dt
import io
import json
import os
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_series as b  # noqa: E402


CACHE_OK = os.path.exists(os.path.join(b.CACHE, "mf.xls")) and os.path.exists(b.DEFAULT_OUT)


@unittest.skipUnless(CACHE_OK, "needs the real download cache")
class UnknownTypeIsAlertNotFailure(unittest.TestCase):
    def test_build_publishes_rest_and_alerts(self):
        from email.utils import parsedate_to_datetime
        meta = json.load(open(os.path.join(b.CACHE, "mf.xls.meta.json"), encoding="utf-8"))
        today = (parsedate_to_datetime(meta["lastModified"]).date() + dt.timedelta(days=13)).isoformat()
        real = b.xls_foreign_codes
        b.xls_foreign_codes = lambda path: real(path) + ["XYZ1129", "XYZ0130"]
        try:
            with tempfile.TemporaryDirectory() as d:
                out = os.path.join(d, "series.json")
                with open(b.DEFAULT_OUT, "rb") as src, open(out, "wb") as dst:
                    dst.write(src.read())
                buf = io.StringIO()
                with redirect_stdout(buf), redirect_stderr(io.StringIO()):
                    rc = b.main(["--offline", "--today", today, "--out", out, "--alerts-out", os.path.join(d, "alerts.out")])
                alerts = open(os.path.join(d, "alerts.out"), encoding="utf-8").read()
                series = json.load(open(out, encoding="utf-8"))["series"]
        finally:
            b.xls_foreign_codes = real
        self.assertEqual(rc, 0)
        self.assertIn("UNKNOWN BOND TYPE XYZ: XYZ0130, XYZ1129 — app needs a code update (BondType.kt)", alerts)
        self.assertIn("::warning::UNKNOWN BOND TYPE XYZ", buf.getvalue())
        self.assertFalse([c for c in series if c.startswith("XYZ")])
        self.assertGreater(len(series), 400)


if __name__ == "__main__":
    unittest.main()
