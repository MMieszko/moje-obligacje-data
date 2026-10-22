#!/usr/bin/env python3
"""Build data/series.json: per-series parameters and rates of Polish retail treasury bonds.

Sources (see tools/data/README.md):
  * MF spreadsheet "Dane dotyczące obligacji detalicznych" (BIFF .xls linked from
    https://www.gov.pl/web/finanse/obligacje-detaliczne1): every series, its sale month, conversion
    price, margin and the rate of every interest period announced so far.
  * Listy emisyjne (PDF, obligacjeskarbowe.pl): early-redemption fee (not in the XLS), first-period
    rate, margin and conversion price (cross-checked against the XLS). Series of the current/next
    month that the XLS does not have yet are built from the LE alone ("source": "LE").
  * Tabele odsetkowe (PDF, obligacjeskarbowe.pl): rolling anchors - each run compares the newest
    announced rates of every type against the official interest table.

Any validation error -> message on stderr, exit code 1, output files untouched.

Usage: python build_series.py [--offline] [--today YYYY-MM-DD] [--out PATH]
Requires: xlrd (BIFF reader), pypdf (PDF text). Both pure Python: pip install xlrd pypdf
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import os
import re
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from email.utils import parsedate_to_datetime

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
CACHE = os.path.join(HERE, "cache")
DEFAULT_OUT = os.path.join(REPO, "data", "series.json")

USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
              "Chrome/130.0 Safari/537.36 MojeObligacje-data/1.0 (+daily check, cached)")
REQUEST_DELAY_S = 1.0
MF_PAGE = "https://www.gov.pl/web/finanse/obligacje-detaliczne1"
OS_BASE = "https://www.obligacjeskarbowe.pl"
# (url, cache name, max age h): the current listing must be fresh daily (new offer appears ~21st-25th)
LE_LISTINGS = [(f"{OS_BASE}/listy-emisyjne/?id=edo1036", "le_listing_0.html", 20),
               (f"{OS_BASE}/archiwum-listow-emisyjnych/", "le_listing_1.html", 24 * 7)]
TO_LISTING = (f"{OS_BASE}/tabela-odsetkowa/", "to/listing.html", 20)

# type -> (months to maturity, months per interest period); mirrors core BondType.kt
TYPES = {
    "OTS": (3, 3), "ROR": (12, 1), "DOR": (24, 1), "TOS": (36, 12),
    "COI": (48, 12), "EDO": (120, 12), "ROS": (72, 12), "ROD": (144, 12),
}
FIXED_RATE = {"OTS", "TOS"}           # one rate for the whole life, XLS has a single column
NO_MARGIN = FIXED_RATE
INDEX_GROUP = {"COI": "CPI", "EDO": "CPI", "ROS": "CPI", "ROD": "CPI", "ROR": "NBP", "DOR": "NBP"}
TO_TYPES = ("ROR", "DOR", "TOS", "COI", "EDO", "ROS", "ROD")   # tabela-odsetkowa has no OTS
TO_ANCHORS_PER_TYPE = 2
RATE_MIN, RATE_MAX = Decimal("0"), Decimal("0.25")
MAX_COUNT_DROP = Decimal("0.05")
SCOPE_GRACE_MONTHS = 1      # keep a series one month after its last maturity month (payout slips to next month)
XLS_MAX_AGE_DAYS = 35       # MF updates the XLS ~9th of each month; > 35 days = an update was missed
MISSING_MONTH_FAIL_DAY = 3  # from this day, no series (XLS or LE) for the current month is an error
PL_MONTHS = {m: i + 1 for i, m in enumerate(
    "stycznia lutego marca kwietnia maja czerwca lipca sierpnia września października listopada grudnia".split())}

# Official values (spec § 7 test vectors and tabele odsetkowe, checked 2026-10-22). Rates in percent.
# "rates" is a prefix: later periods may be added. Long-lived anchors keep every type covered past 2027-02;
# the expiring ones are skipped once out of scope, the rolling tabela-odsetkowa anchors cover the rest.
ANCHORS = {
    "EDO1026": {"rates": "2.50 3.30 3.50 4.40 4.40 7.00 17.60 11.60 5.80 4.40"},
    "COI1026": {"rates": "7.00 11.10 5.30 3.90"},
    "ROS1026": {"rates": "1.50 6.75 17.35 11.35 5.55 4.15"},
    "TOS1026": {"rates": "6.85 6.85 6.85"},
    "EDO1036": {"rate1": "5.35", "fee": "3.00", "margin": "2.00"},
    "TOS1029": {"rates": "4.40 4.40 4.40", "fee": "1.00"},
    "OTS0127": {"rate1": "2.00", "conversionPrice": "100.00"},
    "ROR0127": {"rates": "4.25 4.00 4.00 3.75 3.75 3.75 3.75 3.75 3.75 3.75"},
    "DOR0127": {"rates": "5.90 5.90 5.90 5.90 5.90 5.40 5.40 5.15 5.15 4.90 4.65 4.40 4.15 4.15 4.15 3.90 3.90 3.90 3.90 3.90 3.90 3.90"},
    "COI1027": {"rates": "7.00 5.55 4.15 4.65"},
    "EDO1027": {"rates": "2.70 3.50 4.40 4.40 7.00 17.60 11.60 5.80 4.40 4.90"},
    "ROS1027": {"rates": "1.50 17.35 11.35 5.55 4.15 4.65"},
    "ROD1028": {"rates": "3.00 3.80 4.00 4.90 4.90 7.50 18.10 12.10 6.30 4.90 5.40", "fee": "2.00"},
}
# Static anchors expire with their series; nothing depends on how many are left. The self-renewing
# "deep anchors" (oldest in-scope series of every type, all periods vs interest tables) take over.
MAX_FIRST_RATE_JUMP = Decimal("0.015")   # LE-only series: period-1 rate vs previous month's offer of the type
ALERTS_PATH = os.path.join(HERE, "alerts.out")
UNVERIFIED = "unverified: "              # prefix of verify_to_anchors messages that are alerts, not errors
PDF_LINK = rb"/media_files/[0-9a-f-]+\.pdf"
MARK_MF_PAGE = "obligacji detalicznych".encode()
MARK_LE_LISTING = rb'value="/(?:listy-emisyjne|archiwum-listow-emisyjnych)/\?id='
MARK_TO_LISTING = rb'value="\d+,\d+"'


class BuildError(Exception):
    pass


# ---------------------------------------------------------------- helpers

def add_months(d: dt.date, n: int) -> dt.date:
    y, m = divmod(d.month - 1 + n, 12)
    return dt.date(d.year + y, m + 1, 1) if d.day == 1 else _clamp(d.year + y, m + 1, d.day)


def _clamp(y, m, day):
    while True:
        try:
            return dt.date(y, m, day)
        except ValueError:
            day -= 1


def dec(x, places: int) -> Decimal:
    """XLS float -> Decimal with `places` decimals; refuses values that need more precision."""
    q = Decimal(1).scaleb(-places)
    d = Decimal(repr(float(x))).quantize(q, rounding=ROUND_HALF_UP)
    if abs(Decimal(repr(float(x))) - d) > Decimal("1e-9"):
        raise BuildError(f"value {x!r} needs more than {places} decimals")
    return d


def pl_num(s: str) -> Decimal:
    return Decimal(s.replace(" ", "").replace(",", "."))


def code_maturity_month(code: str, today: dt.date | None = None) -> tuple[int, int]:
    """Series code <TYPE>MMYY -> (year, month) of the last possible maturity (spec § 6).

    Historical types have 1990s codes (IR0995); a year more than 20 years ahead is read as 19YY."""
    mm, yy = int(code[-4:-2]), int(code[-2:])
    year = 2000 + yy
    if year > (today or dt.date.today()).year + 20:
        year -= 100
    return year, mm


def in_scope(code: str, today: dt.date) -> bool:
    """Not yet fully redeemed (+ SCOPE_GRACE_MONTHS): the latest lot matures this month or later."""
    y, m = code_maturity_month(code, today)
    return y * 12 + m - 1 + SCOPE_GRACE_MONTHS >= today.year * 12 + today.month - 1


def series_code(btype: str, sale_month: dt.date) -> str:
    mat = add_months(sale_month.replace(day=1), TYPES[btype][0])
    return f"{btype}{mat.month:02d}{mat.year % 100:02d}"


def expected_known_periods(btype: str, sale_start: dt.date, as_of: dt.date) -> int:
    """Periods already started for a lot bought on the first sale day, as of `as_of`."""
    months, per = TYPES[btype]
    n = 0
    while n < months // per and add_months(sale_start, n * per) <= as_of:
        n += 1
    return n


# ---------------------------------------------------------------- HTTP with cache

_last_request = [0.0]
def _is_html(b: bytes) -> bool:
    return b"<html" in b[:4096].lower() or b"<!doctype html" in b[:4096].lower()


# kind -> check of the downloaded bytes. HTML kinds also require a content marker: a 200 maintenance/WAF
# page is a valid HTML document but not the page we asked for.
SIGNATURES = {"pdf": lambda b: b.startswith(b"%PDF"),
              "xls": lambda b: b.startswith(bytes.fromhex("D0CF11E0A1B11AE1")),
              "html": _is_html,
              "mf_page": lambda b: _is_html(b) and MARK_MF_PAGE in b,
              "le_listing": lambda b: _is_html(b) and re.search(MARK_LE_LISTING, b) is not None,
              "to_listing": lambda b: _is_html(b) and re.search(MARK_TO_LISTING, b) is not None,
              "page_with_pdf": lambda b: _is_html(b) and re.search(PDF_LINK, b) is not None}


def _valid(body: bytes, kind: str | None) -> bool:
    return len(body) > 0 and (kind is None or SIGNATURES[kind](body))


FALLBACKS: list[str] = []   # cache names served stale because the network failed (reported as alerts)


def atomic_write(path: str, body: bytes) -> None:
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=d, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(body)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def fetch(url: str, cache_name: str, max_age_h: float | None, offline: bool, kind: str | None = None,
          referer: str | None = None) -> bytes:
    """GET with file cache. max_age_h=None -> cached forever (immutable documents such as LE PDFs).

    The body must match `kind`'s file signature before it is cached (an HTML error page served with
    200 never lands as a .pdf, a maintenance page never lands as an LE page). A cached file that fails
    the check is ignored. A wrong-content response is not retried (same answer within seconds) and not
    cached, so the next run asks again. If the network fails, a stale but valid cached copy is used with
    a WARNING and recorded in FALLBACKS (freshness rules are checked separately)."""
    path = os.path.join(CACHE, cache_name)
    cached = None
    if os.path.exists(path):
        with open(path, "rb") as f:
            cached = f.read()
        if not _valid(cached, kind):
            print(f"WARNING: cached {cache_name} is not a valid {kind}; ignoring it", file=sys.stderr)
            cached = None
        else:
            age_h = (time.time() - os.path.getmtime(path)) / 3600
            if offline or max_age_h is None or age_h < max_age_h:
                return cached
    if offline:
        raise BuildError(f"--offline and no valid cache for {url} ({cache_name})")
    wait = REQUEST_DELAY_S - (time.time() - _last_request[0])
    if wait > 0:
        time.sleep(wait)
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "pl,en;q=0.8"}
    if referer:
        headers["Referer"] = referer
    req = urllib.request.Request(url, headers=headers)
    body, meta, last_err = None, None, None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                got = r.read()
                status = getattr(r, "status", 200)
                meta = {"url": url, "fetchedAt": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
                        "lastModified": r.headers.get("Last-Modified")}
            if not _valid(got, kind):
                _last_request[0] = time.time()
                raise BuildError(f"{url}: response is not the expected {kind} (HTTP {status}, {len(got)} bytes, "
                                 f"starts {got[:40]!r})")
            body = got
            break
        except BuildError as e:
            if cached is not None:
                print(f"WARNING: {e}; using cached {cache_name}", file=sys.stderr)
                FALLBACKS.append(cache_name)
                return cached
            raise
        except Exception as e:  # noqa: BLE001 - network errors of any kind are retried, then fatal
            last_err = e
            time.sleep(5 * (attempt + 1))
        finally:
            _last_request[0] = time.time()
    if body is None:
        if cached is not None:
            print(f"WARNING: download failed ({last_err}); using cached {cache_name}", file=sys.stderr)
            FALLBACKS.append(cache_name)
            return cached
        raise BuildError(f"download failed: {url}: {last_err}")
    atomic_write(path, body)
    atomic_write(path + ".meta.json", json.dumps(meta).encode("utf-8"))
    return body


def cache_meta(cache_name: str) -> dict:
    try:
        with open(os.path.join(CACHE, cache_name + ".meta.json"), encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}


def pdf_text(path: str) -> str:
    import logging
    import pypdf
    logging.getLogger("pypdf").setLevel(logging.ERROR)   # font-encoding noise; text is validated downstream
    return "\n".join(p.extract_text() or "" for p in pypdf.PdfReader(path).pages)


def cached_pdf_text(pdf_name: str) -> str:
    """Text of a cached PDF, itself cached as .txt. Never leaves an empty/partial .txt behind."""
    pdf_path = os.path.join(CACHE, pdf_name)
    txt_path = pdf_path[:-4] + ".txt"
    if os.path.exists(txt_path) and os.path.getsize(txt_path) > 0:
        with open(txt_path, encoding="utf-8") as f:
            return f.read()
    try:
        text = pdf_text(pdf_path)
    except Exception as e:  # noqa: BLE001 - any pypdf failure is a data error for this document
        raise BuildError(f"cannot read PDF {pdf_name}: {type(e).__name__}: {e}") from e
    if len(text.strip()) < 200:
        raise BuildError(f"PDF {pdf_name}: no text extracted")
    atomic_write(txt_path, text.encode("utf-8"))
    return text


# ---------------------------------------------------------------- parsing: MF page + XLS

def find_xls_url(page_html: str, base: str = MF_PAGE) -> str:
    """Link 'Dane dotyczące obligacji detalicznych' (attachment UUID changes -> never hardcode)."""
    for m in re.finditer(r'<a\b[^>]*href="([^"]+)"[^>]*>(.*?)</a>', page_html, re.S | re.I):
        text = html.unescape(re.sub(r"<[^>]+>", " ", m.group(2)))
        if "Dane dotyczące obligacji detalicznych" in text or "Dane_dotyczace_obligacji_detalicznych" in text.replace("​", ""):
            return urllib.parse.urljoin(base, html.unescape(m.group(1)))
    raise BuildError("MF page: link 'Dane dotyczące obligacji detalicznych' not found")


@dataclass
class XlsSeries:
    code: str
    type: str
    sale_start: dt.date
    rates: list[Decimal]
    margin: Decimal | None
    conversion_price: Decimal | None
    gaps: list[int] = field(default_factory=list)   # 1-based periods blank before a filled one
    source: str = "XLS"                              # "LE" = built from the list emisyjny alone


def _cell_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() in ("", "-"))


def parse_sheet(btype: str, rows: list[list], to_date) -> dict[str, XlsSeries]:
    """One XLS sheet (list of row value lists) -> series. Columns located by header text, not index.

    `to_date` converts the Excel serial in 'Początek sprzedaży' to a date (workbook datemode).
    """
    if not rows:
        raise BuildError(f"sheet {btype}: empty")
    head = [str(v).strip() for v in rows[0]]

    def col(prefix, required=True):
        for i, h in enumerate(head):
            if h.startswith(prefix):
                return i
        if required:
            raise BuildError(f"sheet {btype}: column '{prefix}' not found in header {head}")
        return None

    c_code, c_start, c_conv, c_rate = col("Seria"), col("Początek sprzedaży"), col("Cena zamiany"), col("Oprocentowanie")
    c_int = col("Odsetki", required=False)
    c_margin = col("Marża", required=btype not in NO_MARGIN)
    rate_end = c_int if (c_int is not None and c_int > c_rate) else c_rate + 1
    sub = rows[1] if len(rows) > 1 else []
    has_sub = len(sub) > 0 and _cell_blank(sub[0]) and any(isinstance(v, str) and "w 1." in v for v in sub)
    rate_cols = list(range(c_rate, rate_end))
    if btype not in FIXED_RATE:
        months, per = TYPES[btype]
        if not has_sub or len(rate_cols) != months // per:
            raise BuildError(f"sheet {btype}: expected {months // per} rate columns, got {len(rate_cols)}")
    out = {}
    for row in rows[2 if has_sub else 1:]:
        code = str(row[c_code]).strip().upper() if c_code < len(row) else ""
        if not re.fullmatch(rf"{btype}\d{{4}}", code):
            continue
        vals = [row[c] if c < len(row) else "" for c in rate_cols]
        rates, gaps, seen_blank = [], [], []
        for i, v in enumerate(vals):
            if _cell_blank(v):
                seen_blank.append(i + 1)
                continue
            if seen_blank:
                gaps.extend(seen_blank)
                seen_blank = []
            rates.append(dec(v, 4))
        margin = None
        if c_margin is not None and not _cell_blank(row[c_margin]):
            margin = dec(row[c_margin], 4)
        conv = None if _cell_blank(row[c_conv]) else dec(row[c_conv], 2)
        if code in out:
            raise BuildError(f"sheet {btype}: duplicate series {code}")
        out[code] = XlsSeries(code, btype, to_date(row[c_start]), rates, margin, conv, gaps)
    return out


NON_SERIES_SHEETS = {"Opis", "Dictionary"}


def find_unknown_types(codes, today: dt.date) -> dict[str, list[str]]:
    """Not-yet-redeemed series whose type prefix the engine does not know -> {prefix: [codes]}.

    Fully redeemed historical types (DOS, TOZ, KOS, POS, IR, RS, TZ, SP) drop out via in_scope."""
    out: dict[str, list[str]] = {}
    for code in codes:
        m = re.fullmatch(r"([A-Z]{2,4})(\d{4})", code)
        if m and m.group(1) not in TYPES and 1 <= int(code[-4:-2]) <= 12 and in_scope(code, today):
            out.setdefault(m.group(1), []).append(code)
    return {k: sorted(v) for k, v in sorted(out.items())}


def xls_foreign_codes(path: str) -> list[str]:
    """Series codes from every XLS sheet that is not one of the 8 engine types (and not a doc sheet).

    A new sheet name is itself suspicious, so all its rows are scanned, not only known historical ones."""
    import xlrd
    book = xlrd.open_workbook(path)
    codes = []
    for name in book.sheet_names():
        if name in TYPES or name in NON_SERIES_SHEETS:
            continue
        sh = book.sheet_by_name(name)
        codes += [str(sh.cell_value(r, 0)).strip().upper() for r in range(sh.nrows) if sh.ncols]
    return codes


def parse_xls(path: str, types=tuple(TYPES)) -> dict[str, XlsSeries]:
    import xlrd  # imported lazily so parsing tests of other parts do not need it
    book = xlrd.open_workbook(path)
    to_date = lambda v: xlrd.xldate_as_datetime(float(v), book.datemode).date()  # noqa: E731
    names = set(book.sheet_names())
    out = {}
    for btype in types:
        if btype not in names:
            raise BuildError(f"XLS: sheet {btype} missing (sheets: {sorted(names)})")
        sh = book.sheet_by_name(btype)
        rows = [sh.row_values(r) for r in range(sh.nrows)]
        out.update(parse_sheet(btype, rows, to_date))
    return out


# ---------------------------------------------------------------- parsing: listy emisyjne

def parse_le_listing(page_html: str) -> dict[str, str]:
    """<option value="/listy-emisyjne/?id=edo07829,edo">EDO0829</option> -> {'EDO0829': '/listy-emisyjne/?id=edo07829'}.

    Keyed by the visible label: the site's ids contain typos (edo07829, edo01028). The ',edo' suffix is
    for the page's JS only; requesting it returns 404."""
    out = {}
    for m in re.finditer(r'<option\b[^>]*value="(/(?:listy-emisyjne|archiwum-listow-emisyjnych)/\?id=[^"]+)"[^>]*>(.*?)</option>',
                         page_html, re.S):
        label = re.sub(r"\s+", "", html.unescape(re.sub(r"<[^>]+>", "", m.group(2)))).upper()
        if re.fullmatch(r"[A-Z]{2,4}\d{4}", label):
            out.setdefault(label, html.unescape(m.group(1)).split(",")[0])
    return out


def parse_le_pdf_link(page_html: str) -> str:
    m = re.search(r'(?:href|data|src)="([^"]*/media_files/[0-9a-f-]+\.pdf)', page_html)
    if not m:
        raise BuildError("page: no PDF link")
    return urllib.parse.urljoin(OS_BASE + "/", m.group(1))


@dataclass
class LeTerms:
    fee: Decimal
    first_rate: Decimal | None      # fraction
    conversion_price: Decimal | None
    codes: frozenset = frozenset()  # series codes mentioned in the text (the list's own code must be one)
    margin: Decimal | None = None   # constant margin (periods >= 2 for CPI types), fraction
    sale_month: dt.date | None = None


def normalize_text(t: str) -> str:
    """Collapse whitespace; glue the spaces pypdf inserts before Polish diacritics in some LE PDFs
    ('nie wy ższ ą ni ż 2,00 zł' -> 'nie wyższą niż 2,00 zł'). May glue a word that starts with a
    diacritic to the previous one; none of the phrases matched here starts with one."""
    t = t.replace(" ", " ").replace("­", "")
    t = re.sub(r"\s+", " ", t)
    return re.sub(r"(?<=[A-Za-ząćęłńóśźżĄĆĘŁŃÓŚŹŻ]) (?=[ąćęłńóśźż])", "", t)


def parse_le_text(text: str, btype: str) -> LeTerms:
    t = normalize_text(text)
    amounts = {pl_num(a) for a in re.findall(r"(?:nie wyższą niż|nie wyższa niż|o kwotę)\s+(\d+(?:,\d{2})?)\s*zł", t)}
    if btype == "OTS":
        if amounts or "odsetki od obligacji nie przysługują" not in t:
            raise BuildError(f"OTS LE: unexpected early-redemption wording (amounts {amounts})")
        fee = Decimal("0.00")
    else:
        if len(amounts) != 1:
            raise BuildError(f"LE: early-redemption fee not unique/missing: {sorted(amounts)}")
        fee = amounts.pop().quantize(Decimal("0.01"))
    m = (re.search(r"pierwszym\s+(?:rocznym\s+|miesięcznym\s+)?okresie odsetkowym stopa procentowa wynosi?\s+(\d+,\d+)\s*%", t)
         or re.search(r"Oprocentowanie obligacji (?:jest stałe i )?wynosi\s+(\d+,\d+)\s*%", t))
    first = (pl_num(m.group(1)) / 100).quantize(Decimal("0.0001")) if m else None
    c = re.search(r"Cena zamiany jednej obligacji wynosi\s+(\d+,\d{2})\s*zł", t)
    conv = pl_num(c.group(1)) if c else None
    margins = {pl_num(x) for x in re.findall(r"(?:stałą marżę|marży) w wysokości\s+(\d+,\d+)\s*%", t)}
    margin = (margins.pop() / 100).quantize(Decimal("0.0001")) if len(margins) == 1 else None
    sm = re.search(r"w dniach od \d{1,2} do \d{1,2} ([a-ząćęłńóśźż]+) (\d{4})", t)
    sale = dt.date(int(sm.group(2)), PL_MONTHS[sm.group(1)], 1) if sm and sm.group(1) in PL_MONTHS else None
    codes = frozenset(re.findall(r"(?<![A-Z])([A-Z]{3}\d{4})(?!\d)", t))
    return LeTerms(fee, first, conv, codes, margin, sale)


def le_only_series(code: str, btype: str, le: LeTerms) -> XlsSeries:
    """A series the XLS does not have yet (offer of the current/next month), from its LE alone."""
    months, _ = TYPES[btype]
    if le.sale_month is None:
        raise BuildError(f"{code}: LE-only series without a sale month in the LE text")
    margin = None if btype in NO_MARGIN else le.margin
    return XlsSeries(code, btype, le.sale_month, [le.first_rate] if le.first_rate is not None else [],
                     margin, le.conversion_price, source="LE")


# ---------------------------------------------------------------- parsing: tabela odsetkowa

def parse_to_listing(page_html: str) -> tuple[dict[str, str], dict[str, dict[dt.date, str]]]:
    """tabela-odsetkowa selects -> ({code: issueId}, {issueId: {period start: tableId}})."""
    issues, tables = {}, {}
    for v, t in re.findall(r'<option\b[^>]*value="(\d+),[a-z]{3}"[^>]*>(.*?)</option>', page_html, re.S):
        issues[re.sub(r"\s+", "", html.unescape(t)).upper()] = v
    for tid, iid, t in re.findall(r'<option\b[^>]*value="(\d+),(\d+)"[^>]*>(.*?)</option>', page_html, re.S):
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})\s*-", html.unescape(t))
        if m:
            tables.setdefault(iid, {})[dt.date(*map(int, m.groups()))] = tid
    return issues, tables


def parse_to_text(text: str) -> tuple[Decimal, str, int]:
    """Interest-table PDF -> (rate of the period as fraction, series code, period number)."""
    t = normalize_text(text)
    r = re.search(r"Oprocentowanie w bieżącym okresie:\s*(\d+,\d+)\s*%", t)
    p = re.search(r"Okres odsetkowy:\s*(\d+)", t)
    c = re.search(r"\b([A-Z]{3}\d{4})\b", t)
    if not (r and p and c):
        raise BuildError("interest table: rate/period/code not found")
    return (pl_num(r.group(1)) / 100).quantize(Decimal("0.0001")), c.group(1), int(p.group(1))


def select_to_anchors(scope: dict[str, XlsSeries], as_of: dt.date) -> dict[str, list[tuple[str, int]]]:
    """Per type, candidates in preference order -> {type: [(code, period index k)]}.

    Newest announced period first (period >= 2 preferred, i.e. an index-driven rate); within the same
    start month the newest sale, then the oldest, then the rest."""
    out = {}
    for btype in TO_TYPES:
        per = TYPES[btype][1]
        cands = []
        for x in scope.values():
            if x.type != btype or x.source != "XLS" or not x.rates:
                continue
            k = expected_known_periods(btype, x.sale_start, as_of) - 1 if btype in FIXED_RATE else len(x.rates) - 1
            if k >= 0:
                cands.append(((add_months(x.sale_start, k * per), k >= 1), x.sale_start, x.code, k))
        groups: dict = {}
        for c in cands:
            groups.setdefault(c[0], []).append(c)
        ordered = []
        for key in sorted(groups, reverse=True):
            g = sorted(groups[key], key=lambda c: c[1])
            ordered += [g[-1]] + ([g[0]] if len(g) > 1 else []) + g[1:-1][::-1]
        out[btype] = [(c[2], c[3]) for c in ordered]
    return out


class NoTable(BuildError):
    pass


def load_to_listing(offline: bool):
    url, name, age = TO_LISTING
    return parse_to_listing(fetch(url, name, age, offline, kind="to_listing").decode("utf-8", "replace"))


def check_to_rate(code: str, k: int, start: dt.date, want: Decimal, to_listing, offline: bool) -> None:
    """Official interest table of `code` period k+1 (starting `start`) must show `want`.
    NoTable = the site has no such table; BuildError = the table disagrees or is not that series/period."""
    issues, tables = to_listing
    iid = issues.get(code)
    tid = tables.get(iid, {}).get(start) if iid else None
    if tid is None:
        raise NoTable(f"no interest table for {code} period starting {start}")
    pg = fetch(f"{OS_BASE}/tabela-odsetkowa/?table_id={tid}", f"to/{tid}.html", None, offline,
               kind="page_with_pdf")
    fetch(parse_le_pdf_link(pg.decode("utf-8", "replace")), f"to/{tid}.pdf", None, offline, kind="pdf")
    rate, tcode, period = parse_to_text(cached_pdf_text(f"to/{tid}.pdf"))
    if (tcode, period) != (code, k + 1):
        raise BuildError(f"table {tid} is {tcode} period {period}, expected {code} period {k + 1}")
    if rate != want:
        raise BuildError(f"period {k + 1}: tabela odsetkowa {rate} != data {want}")


def anchor_month(scope, label: str) -> str:
    code, period = label.split("#")
    x = scope[code]
    return add_months(x.sale_start.replace(day=1), (int(period) - 1) * TYPES[x.type][1]).strftime("%Y-%m")


def verify_to_anchors(scope, as_of, offline, per_type: int = TO_ANCHORS_PER_TYPE,
                      max_tries: int = 10, to_listing=None) -> tuple[list[str], list[str]]:
    """Rolling anchors against official interest tables. -> (messages, verified 'CODE#period' labels).

    A table that disagrees is an error. A type whose NEWEST announced period month has no verified table is
    reported with the UNVERIFIED prefix (the caller withholds that month's rates); verifying an older month
    instead does not count, because older months are already locked by the published index."""
    errors, ok = [], []
    to_listing = to_listing or load_to_listing(offline)
    for btype, cands in select_to_anchors(scope, as_of).items():
        if not cands:
            continue
        newest = None
        done, newest_ok = 0, False
        for code, k in cands[:max_tries]:
            if done >= per_type and newest_ok:
                break
            x = scope[code]
            start = add_months(x.sale_start, k * TYPES[x.type][1])
            newest = newest or start
            try:
                check_to_rate(code, k, start, x.rates[0] if x.type in FIXED_RATE else x.rates[k], to_listing, offline)
                ok.append(f"{code}#{k + 1}")
                done += 1
                newest_ok = newest_ok or start == newest
            except NoTable as e:
                print(f"WARNING: rolling anchor {code}: {e}; trying next candidate")
            except BuildError as e:
                if "--offline and no valid cache" in str(e) or "download failed" in str(e):
                    print(f"WARNING: rolling anchor {code}: {e}")
                else:
                    errors.append(f"rolling anchor {code}: {e}")
        if not newest_ok:
            errors.append(f"{UNVERIFIED}{btype}: newest announced period ({newest:%Y-%m}) not verified against "
                          f"tabela odsetkowa (verified only: {', '.join(l for l in ok if l[:3] == btype) or 'none'})")
    return errors, ok


def verify_deep_anchors(scope, offline, to_listing) -> tuple[list[str], list[str], list[str]]:
    """Self-renewing long anchors: per type the oldest series in scope, every period vs its interest table.
    -> (errors, alerts, verified labels). Old-layout PDFs (~before 2020) that do not parse are skipped."""
    errors, alerts, ok = [], [], []
    for btype in TO_TYPES:
        xs = sorted((x for x in scope.values() if x.type == btype and x.source == "XLS" and x.rates),
                    key=lambda x: x.sale_start)
        if not xs:
            continue
        x = xs[0]
        per, got = TYPES[btype][1], 0
        n = 1 if btype in FIXED_RATE else len(x.rates)
        for k in range(n):
            try:
                check_to_rate(x.code, k, add_months(x.sale_start, k * per), x.rates[0] if btype in FIXED_RATE else x.rates[k],
                              to_listing, offline)
                ok.append(f"{x.code}#{k + 1}")
                got += 1
            except NoTable:
                pass
            except BuildError as e:
                if "rate/period/code not found" in str(e) or "--offline and no valid cache" in str(e) or "download failed" in str(e):
                    continue
                errors.append(f"deep anchor {x.code}: {e}")
        if got == 0:
            alerts.append(f"deep anchor {x.code}: no period could be verified against tabela odsetkowa")
    return errors, alerts, ok


# ---------------------------------------------------------------- index consistency (rate - margin)

def implied_index(btype: str, sale_start: dt.date, rates, margin):
    """(group, 'YYYY-MM' period start, rate - margin) for periods >= 2 of CPI/NBP-linked series."""
    if btype not in INDEX_GROUP or margin is None:
        return
    per = TYPES[btype][1]
    for k, r in enumerate(rates):
        if k >= 1:
            yield INDEX_GROUP[btype], add_months(sale_start.replace(day=1), k * per).strftime("%Y-%m"), Decimal(r) - Decimal(margin)


def build_index(series) -> tuple[dict, list[str]]:
    """Index per group and start month; every series starting a period that month must agree."""
    seen: dict = {}
    for x in series:
        for g, mon, v in implied_index(x.type, x.sale_start, x.rates, x.margin):
            seen.setdefault(g, {}).setdefault(mon, {}).setdefault(v, []).append(x.code)
    index, errors = {}, []
    for g, months in seen.items():
        for mon, vals in sorted(months.items()):
            if len(vals) > 1:
                errors.append(f"{g} index {mon}: rate - margin differs across series: "
                              + "; ".join(f"{v} {sorted(c)[:4]}" for v, c in sorted(vals.items())))
            else:
                index.setdefault(g, {})[mon] = next(iter(vals))
    return index, errors


def index_from_published(doc: dict | None) -> dict:
    if not doc:
        return {}
    xs = []
    for code, s in doc.get("series", {}).items():
        if s.get("type") in INDEX_GROUP and s.get("margin") is not None:
            y, m = map(int, s["saleMonth"].split("-"))
            xs.append(XlsSeries(code, s["type"], dt.date(y, m, 1), [Decimal(r) for r in s["rates"]], Decimal(s["margin"]), None))
    return build_index(xs)[0]


_REFERENCE_DOC: list = []


def reference_doc() -> dict | None:
    """The last published data/series.json (validated when it was published)."""
    if not _REFERENCE_DOC:
        doc = None
        if os.path.exists(DEFAULT_OUT):
            with open(DEFAULT_OUT, encoding="utf-8") as f:
                doc = json.load(f)
        _REFERENCE_DOC.append(doc)
    return _REFERENCE_DOC[0]


def reference_index() -> dict:
    return index_from_published(reference_doc())


def offers_from(series_items) -> dict:
    """{(type, 'YYYY-MM' sale month): period-1 rate} from (type, sale month date, rates) triples."""
    return {(t, m.strftime("%Y-%m")): rates[0] for t, m, rates in series_items if rates}


def reference_offers() -> dict:
    return index_offers(reference_doc())


def index_offers(doc: dict | None) -> dict:
    doc = doc or {}
    items = []
    for s in doc.get("series", {}).values():
        y, m = map(int, s["saleMonth"].split("-"))
        items.append((s["type"], dt.date(y, m, 1), [Decimal(r) for r in s["rates"]]))
    return offers_from(items)


# ---------------------------------------------------------------- validation

def pct(s: str) -> Decimal:
    return (Decimal(s) / 100).quantize(Decimal("0.0001"))


def check_anchors(series: dict[str, dict], today: dt.date, withheld=frozenset()) -> tuple[list[str], list[str]]:
    """Static regression anchors, checked on the data BEFORE rate withholding. Out of scope -> skipped;
    series withheld this run -> skipped (alerted elsewhere); missing for any other reason -> error. How many remain does not matter (deep anchors renew themselves)."""
    errors, skipped = [], []
    for code, a in ANCHORS.items():
        if not in_scope(code, today) or code in withheld:
            skipped.append(code)
            continue
        s = series.get(code)
        if s is None:
            errors.append(f"anchor {code}: series missing from output")
            continue
        if "rates" in a:
            want = [str(pct(x)) for x in a["rates"].split()]
            if s["rates"][:len(want)] != want:
                errors.append(f"anchor {code}: rates {s['rates']} do not start with {want}")
        if "rate1" in a and (not s["rates"] or s["rates"][0] != str(pct(a["rate1"]))):
            errors.append(f"anchor {code}: rate1 {s['rates'][:1]} != {pct(a['rate1'])}")
        if "fee" in a and s["earlyRedemptionFee"] != a["fee"]:
            errors.append(f"anchor {code}: fee {s['earlyRedemptionFee']} != {a['fee']}")
        if "margin" in a and s["margin"] != str(pct(a["margin"])):
            errors.append(f"anchor {code}: margin {s['margin']} != {pct(a['margin'])}")
        if "conversionPrice" in a and s["conversionPrice"] != a["conversionPrice"]:
            errors.append(f"anchor {code}: conversionPrice {s['conversionPrice']} != {a['conversionPrice']}")
    return errors, skipped


def validate_series(x: XlsSeries, le: LeTerms | None, as_of: dt.date, index: dict | None = None,
                    offers: dict | None = None) -> list[str]:
    """`index`: {group: {'YYYY-MM': rate - margin}}; default = index of the last published data.
    `offers`: {(type, 'YYYY-MM'): period-1 rate}; default = offers in the last published data. Used for
    LE-only series, whose other checks compare the LE with itself."""
    e = []
    if x.source == "LE" and x.rates:
        offs = reference_offers() if offers is None else offers
        prev = offs.get((x.type, add_months(x.sale_start.replace(day=1), -1).strftime("%Y-%m")))
        if prev is not None and abs(x.rates[0] - prev) > MAX_FIRST_RATE_JUMP:
            e.append(f"{x.code}: LE first rate {x.rates[0]} jumps more than {MAX_FIRST_RATE_JUMP} vs previous month's "
                     f"{x.type} offer {prev} - misread LE?")
    months, per = TYPES[x.type]
    exp_code = add_months(x.sale_start.replace(day=1), months)
    if (exp_code.year % 100, exp_code.month) != (int(x.code[5:7]), int(x.code[3:5])):
        e.append(f"{x.code}: code does not match sale start {x.sale_start} + {months} months")
    if x.gaps:
        e.append(f"{x.code}: gap in rates, blank periods {x.gaps}")
    if not x.rates:
        e.append(f"{x.code}: no rates")
    for r in x.rates:
        if not (RATE_MIN <= r <= RATE_MAX):
            e.append(f"{x.code}: rate {r} outside [{RATE_MIN}, {RATE_MAX}]")
    if x.type in FIXED_RATE:
        if len(x.rates) != 1:
            e.append(f"{x.code}: fixed-rate type expects 1 rate, got {len(x.rates)}")
    else:
        if x.source == "XLS":
            exp = expected_known_periods(x.type, x.sale_start, as_of)
            if not (exp <= len(x.rates) <= min(months // per, exp + 1)):
                e.append(f"{x.code}: {len(x.rates)} rates known, expected {exp} (as of XLS date {as_of}) - stale or missing period")
        if x.margin is None:
            e.append(f"{x.code}: margin missing")
    idx = reference_index() if index is None else index
    for g, mon, v in implied_index(x.type, x.sale_start, x.rates, x.margin):
        ref = idx.get(g, {}).get(mon)
        if ref is not None and ref != v:
            e.append(f"{x.code}: period starting {mon}: rate - margin = {v}, but {g} index for {mon} is {ref}")
    if le is None:
        e.append(f"{x.code}: no list emisyjny (fee unknown)")
    else:
        if x.code not in le.codes:
            e.append(f"{x.code}: list emisyjny text does not mention {x.code} (wrong document?)")
        if le.first_rate is None:
            e.append(f"{x.code}: LE first-period rate not found")
        elif x.rates and le.first_rate != x.rates[0]:
            e.append(f"{x.code}: LE first rate {le.first_rate} != XLS {x.rates[0]}")
        if le.conversion_price != x.conversion_price:
            e.append(f"{x.code}: conversion price LE {le.conversion_price} != XLS {x.conversion_price}")
        if le.margin is not None and x.margin is not None and le.margin != x.margin:
            e.append(f"{x.code}: margin LE {le.margin} != XLS {x.margin}")
        if le.sale_month is not None and le.sale_month != x.sale_start.replace(day=1):
            e.append(f"{x.code}: LE sale month {le.sale_month:%Y-%m} != {x.sale_start:%Y-%m}")
    return e


def freshness_errors(xls_date: dt.date, newest_xls: dt.date, newest_any: dt.date, today: dt.date) -> tuple[list[str], list[str]]:
    """Daily runs: (errors, warnings). XLS updates ~9th; LE of next month ~21st-25th."""
    errors, warnings = [], []
    age = (today - xls_date).days
    if age > XLS_MAX_AGE_DAYS:
        errors.append(f"XLS stale: Last-Modified {xls_date} is {age} days old (> {XLS_MAX_AGE_DAYS})")
    prev_month = add_months(today.replace(day=1), -1)
    if newest_xls.replace(day=1) < prev_month:
        errors.append(f"XLS stale: newest series sold {newest_xls:%Y-%m}, previous month {prev_month:%Y-%m} missing")
    if newest_any.replace(day=1) < today.replace(day=1):
        msg = f"no series (XLS or LE) for {today:%Y-%m}"
        (errors if today.day >= MISSING_MONTH_FAIL_DAY else warnings).append(msg)
    return errors, warnings


def to_json_entry(x: XlsSeries, le: LeTerms | None) -> dict:
    months, per = TYPES[x.type]
    rates = x.rates * (months // per) if x.type in FIXED_RATE else x.rates
    return {
        "type": x.type,
        "saleMonth": x.sale_start.strftime("%Y-%m"),
        "rates": [str(r) for r in rates],
        "margin": None if x.margin is None else str(x.margin),
        "earlyRedemptionFee": None if le is None else str(le.fee),
        "conversionPrice": None if x.conversion_price is None else str(x.conversion_price),
        "firstRateFixed": bool(le and le.first_rate is not None),
        "source": x.source,
    }


# ---------------------------------------------------------------- output: version + manifest

SCHEMA = 1
VOLATILE_KEYS = ("generatedAt", "version", "source")   # ignored when deciding whether content changed


def same_content(prev: dict, new: dict) -> bool:
    strip = lambda d: {k: v for k, v in d.items() if k not in VOLATILE_KEYS}  # noqa: E731
    return strip(prev) == strip(new)


def next_version(prev: int | None, today: dt.date) -> int:
    """YYYYMMDDNN, strictly greater than the previous version (also when that is dated ahead of the clock)."""
    base = int(today.strftime("%Y%m%d")) * 100
    if prev is not None and prev >= base:
        return prev + 1
    return base


def serialize(doc: dict) -> bytes:
    """Canonical published bytes: UTF-8, LF line endings."""
    return (json.dumps(doc, ensure_ascii=False, indent=1) + "\n").encode("utf-8")


def build_manifest(version: int, name: str, body: bytes) -> dict:
    import hashlib
    return {"version": version, "schema": SCHEMA,
            "files": {name: {"sha256": hashlib.sha256(body).hexdigest(), "bytes": len(body)}}}


# ---------------------------------------------------------------- main

def verify_new_months(scope, new_months: dict, verified: dict, offline, to_listing,
                      max_tries: int = 8) -> tuple[list[str], list[str]]:
    """For every index month not yet in the published index (per group), try series that start a period
    (>= 2) in that month until one interest table confirms it. Updates `verified` {group: {months}}.
    -> (errors for tables that disagree, verified labels)."""
    errors, ok = [], []
    for g, months in new_months.items():
        for mon in months:
            if mon in verified.get(g, set()):
                continue
            cands = []
            for x in scope.values():
                if INDEX_GROUP.get(x.type) != g or x.source != "XLS":
                    continue
                per = TYPES[x.type][1]
                for k in range(1, len(x.rates)):
                    if add_months(x.sale_start.replace(day=1), k * per).strftime("%Y-%m") == mon:
                        cands.append((x.sale_start, x.code, k))
            for _, code, k in sorted(cands, reverse=True)[:max_tries]:
                x = scope[code]
                try:
                    check_to_rate(code, k, add_months(x.sale_start, k * TYPES[x.type][1]), x.rates[k], to_listing, offline)
                except NoTable:
                    continue
                except BuildError as e:
                    if not _transient(e):
                        errors.append(f"index month {g} {mon}: {code}: {e}")
                    continue
                ok.append(f"{code}#{k + 1}")
                verified.setdefault(g, set()).add(mon)
                break
    return errors, ok


def le_from_published(code: str, s: dict) -> LeTerms:
    """LE terms of an already published series (they never change), for a day its LE cannot be loaded."""
    y, m = map(int, s["saleMonth"].split("-"))
    return LeTerms(Decimal(s["earlyRedemptionFee"]), Decimal(s["rates"][0]) if s["rates"] else None,
                   None if s["conversionPrice"] is None else Decimal(s["conversionPrice"]), frozenset({code}),
                   None, dt.date(y, m, 1))


def _transient(e: Exception) -> bool:
    return "--offline and no valid cache" in str(e) or "download failed" in str(e)


def report_error(msg: str) -> None:
    print(f"ERROR: {msg}", file=sys.stderr)
    print(f"::error::{msg}")


def write_alerts(alerts: list[str], path: str) -> None:
    """Class (B) problems: data withheld or unverified, the rest still published. One line per alert in
    `path` (the workflow turns a non-empty file into a failed 'alerts' job = e-mail), ::warning:: lines,
    and the GitHub step summary when running in Actions. The file is rewritten on every run."""
    for a in alerts:
        print(f"ALERT: {a}")
        print(f"::warning::{a}")
    atomic_write(path, "".join(a.replace("\n", " ") + "\n" for a in alerts).encode("utf-8"))
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary and alerts:
        with open(summary, "a", encoding="utf-8") as f:
            f.write("### Alerty danych obligacji\n\n" + "".join(f"- {a}\n" for a in alerts))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--offline", action="store_true", help="use only cached downloads")
    ap.add_argument("--today", type=dt.date.fromisoformat, default=dt.date.today(), help="scope date (default: today)")
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--alerts-out", default=ALERTS_PATH)
    ap.add_argument("--xls-max-age-h", type=float, default=12, help="refetch MF page/XLS when cache older")
    args = ap.parse_args(argv)
    FALLBACKS.clear()
    try:
        return build(args)
    except BuildError as e:
        msg = f"FATAL: {e}"
    except Exception as e:  # noqa: BLE001 - unattended run: every failure must be the reported exit-1 path
        msg = f"FATAL: unexpected {type(e).__name__}: {e}"
    print(msg, file=sys.stderr)
    print(f"::error::{msg}")
    return 1


def load_le(code: str, btype: str, listing: dict, offline: bool) -> LeTerms:
    path = listing.get(code)
    if path is None:
        raise BuildError("not in LE listings")
    pg = fetch(OS_BASE + path, f"le/{code}.html", None, offline, kind="page_with_pdf")
    fetch(parse_le_pdf_link(pg.decode("utf-8", "replace")), f"le/{code}.pdf", None, offline, kind="pdf")
    return parse_le_text(cached_pdf_text(f"le/{code}.pdf"), btype)


def build(args) -> int:
    """Two failure classes:
      (A) wrong data (inconsistent, contradicting an official source, unknown type, ...) -> ERROR lines,
          exit 1, nothing written;
      (B) stale / unverifiable / one download failed -> ALERT lines, only the affected series or rates are
          withheld (absent = n/a in the app, never a guess), the rest is published, exit 0."""
    today = args.today
    errors, alerts, warnings = [], [], []
    page = fetch(MF_PAGE, "mf_page.html", args.xls_max_age_h, args.offline, kind="mf_page").decode("utf-8", "replace")
    xls_url = find_xls_url(page)
    fetch(xls_url, "mf.xls", args.xls_max_age_h, args.offline, kind="xls", referer=MF_PAGE)
    lm = cache_meta("mf.xls").get("lastModified")
    if not lm:
        raise BuildError("XLS Last-Modified unknown (needed to check rate completeness)")
    xls_date = parsedate_to_datetime(lm).date()
    xls = parse_xls(os.path.join(CACHE, "mf.xls"))
    scope = {c: s for c, s in xls.items() if in_scope(c, today)}
    if not scope:
        raise BuildError("no series in scope")

    listing = {}
    for url, name, age in LE_LISTINGS:
        body = fetch(url, name, age, args.offline, kind="le_listing").decode("utf-8", "replace")
        for k, v in parse_le_listing(body).items():
            listing.setdefault(k, v)

    unknown = find_unknown_types(set(xls_foreign_codes(os.path.join(CACHE, "mf.xls"))) | set(listing), today)
    for prefix, codes in unknown.items():
        errors.append(f"UNKNOWN BOND TYPE {prefix} (engine knows only {', '.join(TYPES)}): "
                      f"not-yet-redeemed series {', '.join(codes)} -> app code update needed")

    prev = None
    if os.path.exists(args.out):
        with open(args.out, encoding="utf-8") as f:
            prev = json.load(f)
    prev_series = (prev or {}).get("series", {})

    withheld: set[str] = set()
    carried: set[str] = set()
    le_terms = {}
    for code in sorted(scope):
        try:
            le_terms[code] = load_le(code, scope[code].type, listing, args.offline)
        except BuildError as e:
            if code in prev_series:
                # LE terms never change: reuse what was published (validated then) and re-check it vs the XLS
                le_terms[code] = le_from_published(code, prev_series[code])
                carried.add(code)
                alerts.append(f"{code}: list emisyjny unavailable ({e}); published LE terms carried forward")
            else:
                alerts.append(f"{code}: list emisyjny unavailable ({e}); new series withheld")
                withheld.add(code)

    try:
        to_listing = load_to_listing(args.offline)
    except BuildError as e:
        to_listing = None
        alerts.append(f"tabela odsetkowa listing unavailable ({e}); new rates cannot be verified")

    offers = offers_from((x.type, x.sale_start.replace(day=1), x.rates) for x in scope.values())
    for t_m, r in index_offers(prev).items():
        offers.setdefault(t_m, r)

    # offer of the current and next month that the XLS does not have yet -> from the LE alone
    newest_xls = max(x.sale_start for x in scope.values())
    this_month = today.replace(day=1)
    for month in (this_month, add_months(this_month, 1)):
        for btype in TYPES:
            code = series_code(btype, month)
            if code in xls or code not in listing:
                continue
            try:
                le = load_le(code, btype, listing, args.offline)
                x = le_only_series(code, btype, le)
            except BuildError as e:
                alerts.append(f"{code} (new offer, LE-only): {e}; not published yet")
                continue
            verified = False
            if to_listing and x.rates:
                try:
                    check_to_rate(code, 0, x.sale_start, x.rates[0], to_listing, args.offline)
                    verified = True
                except NoTable:
                    pass
                except BuildError as e:
                    if not _transient(e):
                        errors.append(f"{code} (LE-only): {e}")
            if month > this_month and not verified:
                # not on sale yet: wait for the official period-1 table (published ~20th, before the month)
                warnings.append(f"{code} (next month's offer, LE-only): period-1 rate not yet confirmed by "
                                f"tabela odsetkowa; not published yet")
                continue
            prev_offer = offers.get((btype, add_months(month, -1).strftime("%Y-%m")))
            if not verified and prev_offer is None:
                alerts.append(f"{code} (LE-only): period-1 rate {x.rates[:1]} not verifiable (no interest table, "
                              f"no previous {btype} offer); withheld")
                continue
            scope[code] = x
            le_terms[code] = le
            how = "checked vs tabela odsetkowa" if verified else "plausibility vs previous month's offer"
            print(f"LE-only series {code} (sale {month:%Y-%m}, rate 1 {how})")
    newest_any = max(x.sale_start for x in scope.values())

    index, idx_err = build_index([x for x in scope.values() if x.source == "XLS"])
    errors += idx_err
    published = index_from_published(prev)
    for g, months in published.items():
        for mon, v in months.items():
            cur = index.get(g, {}).get(mon)
            if cur is not None and cur != v:
                errors.append(f"{g} index {mon} changed vs published data: {v} -> {cur}")
    merged = {g: {**published.get(g, {}), **index.get(g, {})} for g in set(index) | set(published)}

    entries = {}
    for code in sorted(scope, key=lambda c: (list(TYPES).index(scope[c].type), scope[c].sale_start)):
        x, le = scope[code], le_terms.get(code)
        if le is None:
            continue                                   # withheld above (alerted)
        errors += validate_series(x, le, xls_date, merged, offers)
        entries[code] = to_json_entry(x, le)

    f_err, f_warn = freshness_errors(xls_date, newest_xls, newest_any, today)
    alerts += f_err                                    # stale = (B): nothing wrong is published, it is just old
    warnings += f_warn

    anchor_err, anchor_skip = check_anchors(entries, today, frozenset(withheld))
    errors += anchor_err

    # Every index month that is NOT yet in the published index must be confirmed by an official table
    # (published months are locked: their value may not change). Bootstrap (nothing published): newest only.
    new_months = {g: sorted(m for m in months if m not in published.get(g, {})) for g, months in index.items()}
    if prev is None:
        new_months = {g: [max(months)] for g, months in index.items()}
    to_ok, deep_ok = [], []
    verified_months: dict = {}
    if to_listing:
        to_msgs, to_ok = verify_to_anchors(scope, xls_date, args.offline, to_listing=to_listing)
        errors += [m for m in to_msgs if not m.startswith(UNVERIFIED)]
        warnings += [m[len(UNVERIFIED):] for m in to_msgs if m.startswith(UNVERIFIED)]
        for label in to_ok:
            g = INDEX_GROUP.get(label[:3])
            if g:
                verified_months.setdefault(g, set()).add(anchor_month(scope, label))
        nm_err, nm_ok = verify_new_months(scope, new_months, verified_months, args.offline, to_listing)
        errors += nm_err
        to_ok += nm_ok
        d_err, d_alert, deep_ok = verify_deep_anchors(scope, args.offline, to_listing)
        errors += d_err
        alerts += d_alert
    for g, months in new_months.items():
        missing = [m for m in months if m not in verified_months.get(g, set())]
        if not missing:
            continue
        cutoff = missing[0]
        held = [m for m in months if m >= cutoff]
        alerts.append(f"{g} index: month(s) {', '.join(missing)} not verified against tabela odsetkowa; "
                      f"new rates of periods starting {', '.join(held)} withheld")
        for code, e in entries.items():
            x = scope[code]
            if INDEX_GROUP.get(x.type) != g or x.source != "XLS":
                continue
            per = TYPES[x.type][1]
            for k in range(1, len(x.rates)):
                if add_months(x.sale_start.replace(day=1), k * per).strftime("%Y-%m") in held:
                    e["rates"] = e["rates"][:k]
                    break

    # a series that is published must never vanish while still in scope
    for code, s in prev_series.items():
        if code not in entries and in_scope(code, today):
            entries[code] = dict(s)
            carried.add(code)
            alerts.append(f"{code}: published series missing from today's sources; previous entry carried forward")
    entries = dict(sorted(entries.items(), key=lambda kv: (list(TYPES).index(kv[1]["type"]), kv[1]["saleMonth"])))
    alerts += [f"download failed, stale cache used: {n}" for n in dict.fromkeys(FALLBACKS)]

    prev_n = len(prev.get("series", {})) if prev else None
    if prev_n and len(entries) < prev_n * (1 - MAX_COUNT_DROP):
        errors.append(f"series count dropped {prev_n} -> {len(entries)} (> {MAX_COUNT_DROP:.0%})")

    counts = {t: sum(1 for s in entries.values() if s["type"] == t) for t in TYPES}
    print(f"XLS {xls_url} (Last-Modified {lm}); scope >= {today:%Y-%m} - {SCOPE_GRACE_MONTHS} mo; "
          f"series {len(entries)}: {counts}")
    print(f"rolling anchors verified: {', '.join(to_ok) or 'none'}")
    print(f"deep anchors verified: {len(deep_ok)} periods ({', '.join(sorted({l.split('#')[0] for l in deep_ok}))})")
    if anchor_skip:
        print(f"static anchors skipped (out of scope / withheld): {', '.join(anchor_skip)}")
    for w in warnings:
        print(f"WARNING: {w}")
    write_alerts(alerts, args.alerts_out)
    if errors:
        for e in errors:
            report_error(e)
        report_error(f"FAILED: {len(errors)} validation error(s); {args.out} not written")
        return 1

    now = dt.datetime.now(dt.timezone.utc)
    doc = {
        "schema": SCHEMA,
        "version": 0,
        "generatedAt": now.isoformat(timespec="seconds"),
        "source": (f"MF 'Dane dotyczące obligacji detalicznych' {xls_url} (Last-Modified {xls_date}); "
                   f"early-redemption fee from listy emisyjne, obligacjeskarbowe.pl"),
        "series": entries,
    }
    manifest_path = os.path.join(os.path.dirname(os.path.abspath(args.out)), "manifest.json")
    if prev is not None and same_content(prev, doc):
        version = prev["version"]
        body = serialize(prev)                     # canonical LF bytes of the published content
        with open(args.out, "rb") as f:
            on_disk = f.read()
        if on_disk != body:
            atomic_write(args.out, body)
            print(f"UNCHANGED content (version {version}); rewrote {args.out} in canonical form (line endings)")
        else:
            print(f"UNCHANGED: content equals {args.out} (version {version}), series.json not rewritten")
    else:
        version = next_version(prev.get("version") if prev else None, now.date())
        doc["version"] = version
        body = serialize(doc)
        atomic_write(args.out, body)
        print(f"OK: wrote {args.out} version {version} (previous: {prev_n if prev_n is not None else 'none'})")
    atomic_write(manifest_path, serialize(build_manifest(version, os.path.basename(args.out), body)))
    print(f"OK: wrote {manifest_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
