"""Second audit (2026-10-22, independent reviewer). Each test encodes a property the unattended pipeline SHOULD
have; a failing test documents a finding. Nothing here touches the real download cache or the network.

Run: python -m unittest tools/data/test_audit2.py -v
NOTE: this file is picked up by the workflow's `unittest discover`, so CI fails until the findings are fixed
(or the file is excluded) - intended, same as test_audit.py was.
"""
import datetime as dt
import io
import os
import re
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from decimal import Decimal as D

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_series as b  # noqa: E402

REPO = os.path.dirname(os.path.dirname(HERE))
WORKFLOW = os.path.join(REPO, ".github", "workflows", "data.yml")


def workflow() -> str:
    with open(WORKFLOW, encoding="utf-8") as f:
        return f.read()


class W1_FailedRunNeverSavesCache(unittest.TestCase):
    """actions/cache@v4 saves only when the job succeeds (post-if: success()). (a) First run with an empty cache:
    ~1050 requests / ~330 MB from a US runner; a timeout or a single failed download -> nothing saved -> the next
    day downloads everything again, forever. (b) Combined with W2 it turns one failure into a permanent one."""

    @unittest.skip("workflow-level, see data.yml")
    def test_cache_saved_even_when_build_fails(self):
        wf = workflow()
        saves_always = re.search(r"actions/cache/save@\S+", wf) and re.search(r"if:\s*(always|failure|\!cancelled)\(\)", wf)
        self.assertTrue(saves_always, "cache is only saved on job success (actions/cache post step)")


class W2_TestsGateBuildOnStaleCache(unittest.TestCase):
    """`Testy parsera` runs test_audit.py BEFORE the build. A4 runs a full `--offline` build (default --today) on the
    RESTORED cache, i.e. the one saved by the last successful run. The day the cached XLS turns 36 days old (MF late
    by a few days), A4 fails, the build step - the only thing that would download the new XLS - never runs, the job
    fails, the cache is not saved (W1) -> identical failure every day after MF has published. Deadlock until someone
    deletes the Actions cache by hand. Same for A1 (fixed --today 2026-12-10): it asserts rc != 0, which becomes false
    once the cache holds an XLS from >= 2026-11-05 and the December LEs (~2026-11-25) -> CI red forever from then."""

    @unittest.skip("workflow-level, see data.yml")
    def test_no_offline_full_build_before_the_online_build(self):
        wf = workflow()
        tests_pos, build_pos = wf.find("unittest discover"), wf.find("python tools/data/build_series.py")
        line = wf[tests_pos:wf.find("\n", tests_pos)]
        gated = tests_pos != -1 and tests_pos < build_pos and "test_audit" not in line.replace("test_*.py", "")
        self.assertFalse(gated and "test_*.py" in line,
                         "test_audit.py (offline full builds on the restored cache) runs before, and gates, the online build")

    def test_a1_not_pinned_to_wall_clock(self):
        with open(os.path.join(HERE, "test_audit.py"), encoding="utf-8") as f:
            src = f.read()
        self.assertFalse('"--today", "2026-12-10"' in src,
                         "A1 expects failure on a fixed date using the live cache; it flips to a permanent CI failure "
                         "~2026-11-25 when the cache catches up")


class W3_DeployOnlyWhenCommitted(unittest.TestCase):
    """deploy runs only if this run committed data/. If the push succeeds and deploy-pages fails (transient), every
    later run sees no git diff -> changed=false -> Pages keeps serving the old manifest indefinitely (one email)."""

    @unittest.skip("workflow-level, see data.yml")
    def test_deploy_not_gated_on_this_runs_commit(self):
        m = re.search(r"deploy:.*?\n\s+if:\s*([^\n]+)", workflow(), re.S)
        cond = m.group(1) if m else ""
        self.assertNotIn("changed == 'true'", cond, f"deploy condition: {cond}")


class W4_UnpinnedParsers(unittest.TestCase):
    """`pip install xlrd pypdf` unpinned: a new pypdf changes text extraction (spaces, diacritics) under regexes that
    were tuned on 6.21.2 -> spurious failures, or a different number matched."""

    def test_dependencies_pinned(self):
        m = re.search(r"pip install ([^\n]+)", workflow())
        self.assertTrue(m and "pypdf==" in m.group(1) and "xlrd==" in m.group(1), m.group(0) if m else "no pip line")


class R1_NewestMonthMustBeVerified(unittest.TestCase):
    """Rolling anchors: a missing table for the newest period silently falls back to an OLDER month. A type can
    then be 'verified' while the rates just added (the only ones not already locked by the published index) are not.
    Today this is caught only by accident (max_tries=6 exhausted for EDO/ROS/ROD); COI passes on September tables
    with every 2026-10 COI rate shifted by +0.10 pp (reproduced on the real cache)."""

    def test_fallback_to_older_month_is_not_verification(self):
        scope = {"EDO1035": b.XlsSeries("EDO1035", "EDO", dt.date(2025, 10, 1), [D("0.0560"), D("0.0550")], D("0.0200"), D("99.90")),
                 "EDO0935": b.XlsSeries("EDO0935", "EDO", dt.date(2025, 9, 1), [D("0.0560"), D("0.0540")], D("0.0200"), D("99.90"))}
        listing = ('<option value="1,edo">EDO1035</option><option value="2,edo">EDO0935</option>'
                   '<option value="11,2">2026-09-01 - 2027-09-01</option>').encode()

        def fake_fetch(url, name, age, offline, kind=None, referer=None):
            if name == "to/listing.html":
                return listing
            if name.endswith(".html"):
                return b'<html><a href="/media_files/0123abcd-0000.pdf">pdf</a></html>'
            return b"%PDF-1.4"

        real = b.fetch, b.cached_pdf_text
        b.fetch = fake_fetch
        b.cached_pdf_text = lambda n: "EDO0935 Okres odsetkowy: 2 Oprocentowanie w bieżącym okresie: 5,40 %"
        try:
            with redirect_stdout(io.StringIO()):
                errors, ok = b.verify_to_anchors(scope, dt.date(2026, 10, 9), True, per_type=1)
        finally:
            b.fetch, b.cached_pdf_text = real
        self.assertEqual(ok, ["EDO0935#2"])
        self.assertTrue([e for e in errors if "EDO" in e],
                        "EDO counted as verified by a 2026-09 table while its 2026-10 rate (newest) was never checked")


class R2_LeOnlyChecksAreTautological(unittest.TestCase):
    """A LE-only series is validated against the LE it was built from (rate 1, margin, conversion price, sale month
    all come from the same text). A misread first rate that is in range passes; nothing independent is consulted
    (the tabela-odsetkowa listing already has period-1 tables of the current month's series, e.g. TOS1029, ROR1027)."""

    def test_misread_first_rate_detected(self):
        with open(os.path.join(HERE, "testdata", "le_ror1027.txt"), encoding="utf-8") as f:
            t = b.parse_le_text(f.read(), "ROR")
        wrong = b.LeTerms(t.fee, D("0.0040"), t.conversion_price, t.codes, t.margin, t.sale_month)   # 0,40% instead of 4,00%
        x = b.le_only_series("ROR1027", "ROR", wrong)
        self.assertTrue(b.validate_series(x, wrong, dt.date(2026, 10, 9), {}, {("ROR", "2026-09"): D("0.0400")}),  # fixture: ROR0927 offer 4,00%
                        "LE-only series with a misread (in-range) first rate passes validation")


class C1_HtmlErrorPageCachedForever(unittest.TestCase):
    """The 'html' signature (<html in the first 4 KB) accepts any HTML page, including a 200 maintenance/WAF page.
    LE and tabela-odsetkowa pages are cached with max_age=None, so such a page is reused forever and every later
    ONLINE run fails with 'page: no PDF link'. (In CI it self-heals only because failed jobs never save the cache -
    fixing W1 without this makes it permanent.)"""

    def test_bad_page_is_refetched(self):
        pages = [b"<html><body>Przerwa techniczna</body></html>",
                 b'<html><a href="/media_files/0123abcd-0000.pdf">LE</a></html>', b"%PDF-1.4 x"]

        class Resp:
            def __init__(self, body):
                self.body, self.headers = body, {}

            def __enter__(self):
                return self

            def __exit__(self, *a):
                return False

            def read(self):
                return self.body

        served = []

        def fake_urlopen(req, timeout=None):
            if not served:
                body = pages[0]                                   # first request: maintenance page, HTTP 200
            else:
                body = pages[2] if "/media_files/" in req.full_url else pages[1]
            served.append(req.full_url)
            return Resp(body)

        with open(os.path.join(HERE, "testdata", "le_ror1027.txt"), encoding="utf-8") as f:
            le_text = f.read()
        old = b.CACHE, b.urllib.request.urlopen, b.cached_pdf_text, b.REQUEST_DELAY_S
        with tempfile.TemporaryDirectory() as d:
            b.CACHE, b.urllib.request.urlopen, b.REQUEST_DELAY_S = d, fake_urlopen, 0
            b.cached_pdf_text = lambda n: le_text
            listing = {"ROR1027": "/listy-emisyjne/?id=ror1027"}
            try:
                with self.assertRaises(b.BuildError):
                    b.load_le("ROR1027", "ROR", listing, False)        # day 1: maintenance page
                try:
                    b.load_le("ROR1027", "ROR", listing, False)        # day 2: site is fine again
                except b.BuildError as e:
                    self.fail(f"cached maintenance page reused on the next online run: {e}")
            finally:
                b.CACHE, b.urllib.request.urlopen, b.cached_pdf_text, b.REQUEST_DELAY_S = old


if __name__ == "__main__":
    unittest.main()
