"""Third audit (2026-10-22, independent reviewer). Failing test = finding. Full offline builds on the real cache
(read-only: the cache is never written by --offline runs; output goes to a temp dir). Local only, not CI.

Run: python -m unittest tools/data/test_audit3.py -v
"""
import datetime as dt
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_series as b  # noqa: E402

REAL_CACHE = os.path.join(HERE, "cache")
REPO_SERIES = os.path.join(os.path.dirname(os.path.dirname(HERE)), "data", "series.json")
HAVE_CACHE = os.path.exists(os.path.join(REAL_CACHE, "mf.xls")) and os.path.exists(REPO_SERIES)
TODAY = "2026-10-22"


def run(d, prev_doc, patches=None):
    """Offline build into temp dir d with prev_doc as the published series.json. -> (rc, doc|None, alerts, stdout)"""
    out = os.path.join(d, "series.json")
    with open(out, "wb") as f:
        f.write(b.serialize(prev_doc))
    saved = {k: getattr(b, k) for k in (patches or {})}
    for k, v in (patches or {}).items():
        setattr(b, k, v)
    o, e = io.StringIO(), io.StringIO()
    try:
        with redirect_stdout(o), redirect_stderr(e):
            rc = b.main(["--offline", "--today", TODAY, "--out", out, "--alerts-out", os.path.join(d, "alerts.out")])
    finally:
        for k, v in saved.items():
            setattr(b, k, v)
    doc = json.load(open(out, encoding="utf-8"))
    alerts = open(os.path.join(d, "alerts.out"), encoding="utf-8").read()
    return rc, doc, alerts, o.getvalue() + e.getvalue()


def prev():
    return json.load(open(REPO_SERIES, encoding="utf-8"))


def lost_rates(old, new):
    """codes whose published rates got shorter (a value the app showed yesterday is gone today)."""
    return sorted(c for c, s in old["series"].items()
                  if c in new["series"] and len(new["series"][c]["rates"]) < len(s["rates"]))


@unittest.skipUnless(HAVE_CACHE, "needs real cache")
class F1_WithholdingRemovesPublishedRates(unittest.TestCase):
    """R1 withholding truncates the newest month's rates even when yesterday's published file already had them
    (verified then). Trigger: interest-table listing unavailable (no cache + network failure)."""

    def test_published_rate_never_withdrawn(self):
        def boom(offline):
            raise b.BuildError("download failed: simulated")
        with tempfile.TemporaryDirectory() as d:
            p = prev()
            rc, doc, alerts, log = run(d, p, {"load_to_listing": boom})
            lost = lost_rates(p, doc)
            self.assertEqual(lost, [], f"rc={rc}, version {p['version']} -> {doc['version']}; {len(lost)} series lost "
                                       f"a published rate, e.g. {lost[:5]}; alerts: {alerts[:300]}")


@unittest.skipUnless(HAVE_CACHE, "needs real cache")
class F2_LeFailureDropsPublishedSeries(unittest.TestCase):
    """A failed/missing LE for a series that is already published removes the WHOLE series (exit 0, alert):
    the app loses the series of a user's lot. Below the 5 % count-drop guard (~22 series) nothing stops it."""

    def test_published_series_not_dropped(self):
        real = b.load_le

        def flaky(code, btype, listing, offline):
            if code in ("EDO1030", "ROR1126"):
                raise b.BuildError("not in LE listings")
            return real(code, btype, listing, offline)
        with tempfile.TemporaryDirectory() as d:
            p = prev()
            rc, doc, alerts, log = run(d, p, {"load_le": flaky})
            gone = sorted(set(p["series"]) & {"EDO1030", "ROR1126"} - set(doc["series"]))
            self.assertFalse(rc == 0 and gone, f"rc=0 and published series removed: {gone}")


@unittest.skipUnless(HAVE_CACHE, "needs real cache")
class F3_UnlockedOlderMonthPublishedUnverified(unittest.TestCase):
    """R1 only verifies the NEWEST index month. A month that was withheld yesterday (so it is NOT locked by the
    published index) is published unverified as soon as a newer month exists. Here: published file without the
    2026-09 CPI periods (as if withheld), XLS with every 2026-09 CPI rate +0.10 pp (consistent across series)."""

    def test_wrong_unlocked_month_rejected(self):
        real = b.parse_xls
        bump = D("0.0010")

        def shifted(path, types=tuple(b.TYPES)):
            xs = real(path, types)
            for x in xs.values():
                if b.INDEX_GROUP.get(x.type) == "CPI":
                    per = b.TYPES[x.type][1]
                    x.rates = [r + bump if k >= 1 and b.add_months(x.sale_start.replace(day=1), k * per).strftime("%Y-%m") == "2026-09"
                               else r for k, r in enumerate(x.rates)]
            return xs
        p = prev()
        hit = []
        for c, s in p["series"].items():
            if s["type"] in ("COI", "EDO", "ROS", "ROD"):
                y, m = map(int, s["saleMonth"].split("-"))
                for k in range(1, len(s["rates"])):
                    if b.add_months(dt.date(y, m, 1), k * 12).strftime("%Y-%m") == "2026-09":
                        s["rates"] = s["rates"][:k]
                        hit.append(c)
                        break
        self.assertTrue(hit)
        with tempfile.TemporaryDirectory() as d:
            rc, doc, alerts, log = run(d, p, {"parse_xls": shifted})
            wrong = [c for c in hit if len(doc["series"][c]["rates"]) > len(p["series"][c]["rates"])]
            self.assertFalse(rc == 0 and wrong, f"rc=0, {len(wrong)} series published with an unverified, wrong 2026-09 "
                                                f"rate, e.g. {wrong[:4]}; alerts: {alerts[:200]!r}")


if __name__ == "__main__":
    unittest.main()
