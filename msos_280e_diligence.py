#!/usr/bin/env python3
"""
MSOS Section 280E Diligence Pipeline
====================================

Quantifies the IRC Section 280E tax penalty embedded in the audited financial
statements of the largest U.S. multi-state cannabis operators (MSOs) and
restates their net income pro forma for Schedule III rescheduling.

Pipeline
    1. Target identification   AdvisorShares Pure US Cannabis ETF (MSOS) holdings
                               file. Total-return-swap and direct-equity lines are
                               aggregated by issuer, cash/collateral is excluded.
    2. Ticker / CIK resolution SEC company_tickers.json (OTC tickers such as
                               CURLF, GTBIF, TCNNF, VRNOF).
    3. Primary-source filing   EDGAR submissions API -> most recent audited annual
                               report: Form 10-K, or Form 40-F for the Canadian-
                               domiciled MSOs (Curaleaf, Cresco, Glass House).
    4. Footnote isolation      Inline XBRL `us-gaap:IncomeTaxDisclosureTextBlock`
                               (heading heuristic as fallback) plus the
                               consolidated statement of operations.
    5. LLM parsing             Claude, constrained by a strict JSON schema,
                               transcribes every line of the statutory-to-
                               effective rate reconciliation. Deterministic code,
                               not the LLM, selects the 280E line(s), applies the
                               unit scale and checks that the reconciliation ties.
    6. Pro forma               Pro Forma Net Income = Reported Net Income
                                                      + Section 280E penalty.
    7. Validation / output     Net income cross-checked against the issuer's XBRL
                               company facts; JSON / CSV / Markdown report with an
                               executive summary generated from the computed
                               figures and linked to every source filing.

Environment
    MSOS_ANTHROPIC_API_KEY  Anthropic API key for the extraction calls (falls back to
                       ANTHROPIC_API_KEY). The project-specific name avoids colliding
                       with Claude Code's own credentials when run in a cloud session.
                       Never hard-code it.
    SEC_USER_AGENT     "Firm Name contact@firm.com" - required by the SEC's
                       fair-access policy (https://www.sec.gov/os/accessing-edgar-data).

Usage
    python msos_280e_diligence.py                 # full run (top 5 operators)
    python msos_280e_diligence.py --dry-run       # everything except the LLM call
    python msos_280e_diligence.py --tickers CURLF,TCNNF,GTBIF,VRNOF,CRLBF
"""

from __future__ import annotations

import argparse
import csv
import difflib
import hashlib
import io
import json
import logging
import os
import random
import re
import sys
import time
import warnings
from collections import Counter
from dataclasses import asdict, dataclass, field
from datetime import date, datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any, Literal, Optional

import requests
from bs4 import BeautifulSoup, Comment, NavigableString, Tag, XMLParsedAsHTMLWarning
from pydantic import BaseModel, ConfigDict, ValidationError

try:  # The LLM stage is optional (--dry-run), so the SDK is imported lazily.
    import anthropic
except ImportError:  # pragma: no cover
    anthropic = None  # type: ignore[assignment]

try:  # Optional convenience for local runs: load secrets from an untracked .env file.
    from dotenv import load_dotenv
except ImportError:  # pragma: no cover
    load_dotenv = None  # type: ignore[assignment]

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)

log = logging.getLogger("msos_280e")

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

MSOS_HOLDINGS_URL = (
    "https://advisorshares.com/wp-content/uploads/csv/holdings/"
    "AdvisorShares_MSOS_Holdings_File.xlsx"
)
MSOS_FUND_PAGE = "https://advisorshares.com/etfs/msos/"
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik:010d}.json"
SEC_COMPANYFACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{acc}/{doc}"

# Audited annual reports. Canadian-domiciled MSOs file Form 40-F rather than 10-K;
# their U.S. GAAP financial statements carry the same income-tax footnote.
ANNUAL_FORMS = ("10-K", "40-F", "20-F")

DEFAULT_MODEL = "claude-opus-5-5"
DEFAULT_EFFORT = "high"
PROMPT_VERSION = "2026-09-30.3"  # bump to invalidate cached LLM extractions

# Browser-style UA for the fund sponsor's public website (SEC gets its own UA).
WEB_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/128.0 Safari/537.36"
)

# Last-resort CIK map, used only when company_tickers.json cannot resolve an
# issuer. Each CIK was verified against the issuer's EDGAR filing index.
SEED_CIKS: dict[str, int] = {
    "curaleaf": 1756770,
    "trulieve cannabis": 1754195,
    "green thumb industries": 1795139,
    "verano": 1848416,
    "cresco labs": 1832928,
    "glass house brands": 1848731,
    "terrascend": 1778129,
}

# Holdings-file rows that are not operating-company equity exposure.
NON_EQUITY_RE = re.compile(
    r"collateral|\bcash\b|money\s*market|treasury|t-?bill|\bbills?\b|government|"
    r"liquidity|repurchase|\brepo\b|dreyfus|federated|fidelity|blackrock|"
    r"goldman\s+sachs\s+fin|first\s+american|us\s+dollar|other\s+assets|"
    r"net\s+assets|liabilities|\bfund\b|\betf\b|receivable|payable",
    re.I,
)

# Tokens dropped when normalising issuer names for matching.
NAME_STOPWORDS = {
    "swap", "swaps", "trs", "total", "return", "inc", "incorporated", "corp",
    "corporation", "co", "company", "ltd", "limited", "plc", "llc", "lp", "sa",
    "nv", "holdings", "holding", "group", "the", "sub", "subordinate", "voting",
    "restricted", "shares", "share", "class", "cl", "a", "b", "common", "stock",
    "ordinary", "svs", "mvs", "new", "reg", "com", "npv", "cad", "usd", "otc",
    "cn", "us", "pv", "ord", "equity", "and",
}


# --------------------------------------------------------------------------- #
# Errors
# --------------------------------------------------------------------------- #


class PipelineError(RuntimeError):
    """A company-level failure: recorded in the report, the run continues."""


class FatalError(RuntimeError):
    """A run-level failure (bad credentials, bad configuration): abort."""


class HttpNotFound(PipelineError):
    pass


# --------------------------------------------------------------------------- #
# HTTP layer: throttling, retries, SEC fair-access handling, on-disk cache
# --------------------------------------------------------------------------- #


def _parse_retry_after(value: str | None) -> float | None:
    if not value:
        return None
    try:
        return max(0.0, float(value))
    except ValueError:
        try:
            dt = parsedate_to_datetime(value)
            return max(0.0, (dt - datetime.now(dt.tzinfo)).total_seconds())
        except (TypeError, ValueError):
            return None


class HttpClient:
    """GET-only HTTP client with a request-rate ceiling, exponential backoff and
    a content-addressed disk cache.

    SEC specifics (https://www.sec.gov/os/accessing-edgar-data):
      * max 10 requests/second per client - we default to 5;
      * a declared User-Agent with contact e-mail is mandatory;
      * exceeding the rate returns 403 "Request Rate Threshold Exceeded" (or 429)
        and the client is blocked for roughly ten minutes, so we back off hard and
        halve our own request rate instead of hammering the endpoint.
    """

    RETRYABLE_STATUS = {408, 425, 429, 500, 502, 503, 504}
    RATE_LIMIT_MARKERS = ("request rate threshold exceeded", "rate limit", "too many requests")
    UNDECLARED_UA_MARKERS = ("undeclared automated tool", "declare your traffic")

    def __init__(
        self,
        user_agent: str,
        *,
        max_rps: float,
        cache_dir: Path | None,
        timeout: float = 45.0,
        max_retries: int = 6,
        base_backoff: float = 2.0,
        rate_limit_cooldown: float = 60.0,
        max_backoff: float = 600.0,
        session: requests.Session | None = None,
        sleep=time.sleep,
    ) -> None:
        self.session = session or requests.Session()
        self.session.headers.update({"User-Agent": user_agent, "Accept-Encoding": "gzip, deflate"})
        self.min_interval = 1.0 / max_rps
        self.cache_dir = cache_dir
        self.timeout = timeout
        self.max_retries = max_retries
        self.base_backoff = base_backoff
        self.rate_limit_cooldown = rate_limit_cooldown
        self.max_backoff = max_backoff
        self._sleep = sleep
        self._last_request = 0.0
        if cache_dir:
            cache_dir.mkdir(parents=True, exist_ok=True)

    # -- cache ---------------------------------------------------------------
    def _cache_path(self, url: str) -> Path | None:
        if not self.cache_dir:
            return None
        return self.cache_dir / (hashlib.sha256(url.encode()).hexdigest()[:40] + ".bin")

    def _cache_read(self, url: str, ttl: float | None) -> bytes | None:
        path = self._cache_path(url)
        if not path or not path.exists():
            return None
        if ttl is not None and time.time() - path.stat().st_mtime > ttl:
            return None
        return path.read_bytes()

    def _cache_write(self, url: str, content: bytes) -> None:
        path = self._cache_path(url)
        if path:
            tmp = path.with_suffix(".tmp")
            tmp.write_bytes(content)
            tmp.replace(path)

    # -- request loop ----------------------------------------------------------
    def _throttle(self) -> None:
        wait = self._last_request + self.min_interval - time.monotonic()
        if wait > 0:
            self._sleep(wait)
        self._last_request = time.monotonic()

    def _backoff(self, attempt: int) -> float:
        return min(self.max_backoff, self.base_backoff * (2 ** attempt) + random.uniform(0, 1))

    def get(self, url: str, *, ttl: float | None = None, use_cache: bool = True) -> bytes:
        """Fetch `url`. `ttl=None` caches forever (EDGAR archives are immutable);
        pass a TTL in seconds for mutable endpoints such as submissions JSON."""
        if use_cache:
            cached = self._cache_read(url, ttl)
            if cached is not None:
                log.debug("cache hit %s", url)
                return cached

        last_problem = "unknown error"
        for attempt in range(self.max_retries + 1):
            self._throttle()
            try:
                resp = self.session.get(url, timeout=self.timeout)
            except (requests.exceptions.ProxyError, requests.exceptions.SSLError) as exc:
                # Policy denials and certificate failures do not heal on retry.
                raise FatalError(
                    f"Connection to {url.split('/')[2]} refused before reaching the server ({type(exc).__name__}). "
                    "Check that your network/proxy policy allows this host and that the CA bundle is trusted."
                ) from exc
            except (requests.ConnectionError, requests.Timeout) as exc:
                last_problem = f"{type(exc).__name__}: {exc}"
                delay = self._backoff(attempt)
            else:
                if resp.status_code == 200:
                    self._cache_write(url, resp.content)
                    return resp.content
                if resp.status_code == 404:
                    raise HttpNotFound(f"404 Not Found: {url}")
                body = resp.text[:2000].lower() if resp.content else ""
                retry_after = _parse_retry_after(resp.headers.get("Retry-After"))
                last_problem = f"HTTP {resp.status_code}"
                if resp.status_code == 403 and any(m in body for m in self.UNDECLARED_UA_MARKERS):
                    raise FatalError(
                        "SEC rejected the request as an undeclared automated tool. Set "
                        "SEC_USER_AGENT to 'Firm Name contact@firm.com' (SEC fair-access policy)."
                    )
                if resp.status_code == 429 or (
                    resp.status_code == 403 and any(m in body for m in self.RATE_LIMIT_MARKERS)
                ):
                    # Rate-limited: slow down for the rest of the run and cool off.
                    self.min_interval = min(self.min_interval * 2, 5.0)
                    delay = max(retry_after or 0.0, self.rate_limit_cooldown * (attempt + 1))
                    delay = min(delay, self.max_backoff)
                    log.warning(
                        "Rate limited by %s (HTTP %s); now at %.1f req/s, cooling off %.0fs",
                        url.split("/")[2], resp.status_code, 1 / self.min_interval, delay,
                    )
                elif resp.status_code in self.RETRYABLE_STATUS:
                    delay = retry_after if retry_after is not None else self._backoff(attempt)
                else:
                    raise PipelineError(f"HTTP {resp.status_code} for {url}: {resp.text[:300]!r}")

            if attempt == self.max_retries:
                break
            log.info("GET %s failed (%s); retry %d/%d in %.1fs",
                     url, last_problem, attempt + 1, self.max_retries, delay)
            self._sleep(delay)
        raise PipelineError(f"GET {url} failed after {self.max_retries + 1} attempts: {last_problem}")

    def get_json(self, url: str, *, ttl: float | None = None, use_cache: bool = True) -> Any:
        return json.loads(self.get(url, ttl=ttl, use_cache=use_cache))


# --------------------------------------------------------------------------- #
# Stage 1 - MSOS holdings
# --------------------------------------------------------------------------- #


@dataclass
class Holding:
    name: str
    ticker: str | None
    weight_pct: float | None
    market_value: float | None
    is_swap: bool


@dataclass
class Issuer:
    key: str
    name: str
    weight_pct: float
    market_value: float
    tickers: list[str]
    components: list[str]
    cik: int | None = None
    sec_ticker: str | None = None
    sec_name: str | None = None
    status: str = "pending"


def issuer_key(name: str) -> str:
    """Normalise an issuer or security name so that 'Curaleaf Holdings Inc Swap',
    'CURALEAF HOLDINGS INC' and SEC's 'Curaleaf Holdings, Inc.' share one key."""
    s = name.lower().replace("&", " and ")
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    tokens = [t for t in s.split() if t not in NAME_STOPWORDS and not re.search(r"\d", t)]
    return " ".join(tokens)


def _to_float(value: Any) -> tuple[float | None, bool]:
    """Parse numbers like '9.72%', '$1,234.5', '(12.0)'. Returns (value, had_percent_sign)."""
    if value is None:
        return None, False
    if isinstance(value, (int, float)):
        return float(value), False
    s = str(value).strip()
    if not s or s in {"-", "—", "–", "N/A", "n/a"}:
        return None, False
    pct = "%" in s
    neg = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[,$%()\s]", "", s)
    try:
        v = float(s)
    except ValueError:
        return None, pct
    return (-v if neg else v), pct


_HEADER_PATTERNS = {
    "name": re.compile(r"(security\s*)?(name|description)|^holding|^security$|^issuer", re.I),
    "ticker": re.compile(r"ticker|symbol", re.I),
    "weight": re.compile(r"weight|%\s*of\s*(net\s*)?(assets|fund|nav)|percent|%\s*nav|^%$|pct", re.I),
    "value": re.compile(r"market\s*value|notional|exposure|^value", re.I),
    "type": re.compile(r"(security|asset)\s*type|^type$|instrument", re.I),
    "date": re.compile(r"^(as\s*of\s*)?date$|as\s*of", re.I),
}


def _rows_from_bytes(data: bytes, filename: str) -> list[list[Any]]:
    if data[:2] == b"PK" or filename.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl  # local import: only needed for the holdings file

        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        rows: list[list[Any]] = []
        for ws in wb.worksheets:  # concatenate sheets; header detection finds the right block
            rows.extend([list(r) for r in ws.iter_rows(values_only=True)])
        return rows
    text = data.decode("utf-8-sig", errors="replace")
    return [list(r) for r in csv.reader(io.StringIO(text))]


def parse_holdings(data: bytes, filename: str = "holdings.xlsx") -> tuple[list[Holding], str | None]:
    """Parse the AdvisorShares holdings file (XLSX or CSV) without assuming a fixed
    layout: locate the header row by column semantics, then read the rows below."""
    rows = _rows_from_bytes(data, filename)
    header_idx, cols = None, {}
    for i, row in enumerate(rows[:60]):
        cells = [str(c).strip() if c is not None else "" for c in row]
        found: dict[str, int] = {}
        for j, cell in enumerate(cells):
            if not cell or len(cell) > 60:
                continue
            for key, pat in _HEADER_PATTERNS.items():
                if key not in found and pat.search(cell):
                    if key == "name" and re.search(r"ticker|symbol|fund|account|portfolio", cell, re.I):
                        continue
                    found[key] = j
                    break
        if "name" in found and ("weight" in found or "value" in found):
            header_idx, cols = i, found
            break
    if header_idx is None:
        raise PipelineError("Could not locate a header row (name + weight/market value) in holdings file")

    as_of: str | None = None
    for row in rows[:header_idx]:  # 'As of 09/29/2026' banner above the table
        for c in row:
            m = re.search(r"as\s+of[:\s]+([0-9/\-]+)", str(c or ""), re.I)
            if m:
                as_of = m.group(1)

    def cell(row: list[Any], key: str) -> Any:
        j = cols.get(key)
        return row[j] if j is not None and j < len(row) else None

    raw: list[tuple[str, str | None, float | None, bool, float | None, bool]] = []
    for row in rows[header_idx + 1:]:
        name = str(cell(row, "name") or "").strip()
        if not name:
            continue
        if as_of is None and cell(row, "date") is not None:
            d = cell(row, "date")
            as_of = d.date().isoformat() if isinstance(d, datetime) else str(d)
        weight, weight_pct_sign = _to_float(cell(row, "weight"))
        value, _ = _to_float(cell(row, "value"))
        sec_type = str(cell(row, "type") or "")
        is_swap = bool(re.search(r"swap|\btrs\b", f"{name} {sec_type}", re.I))
        ticker = str(cell(row, "ticker")).strip() if cell(row, "ticker") not in (None, "") else None
        raw.append((name, ticker, weight, weight_pct_sign, value, is_swap))

    weights = [abs(w) for _, _, w, _, _, _ in raw if w is not None]
    any_pct_sign = any(p for _, _, _, p, _, _ in raw)
    # Weights are either fractions (0.0972) or percentages (9.72 / '9.72%').
    scale = 100.0 if (weights and not any_pct_sign and max(weights) <= 1.5) else 1.0

    holdings = [
        Holding(name=n, ticker=t, weight_pct=(w * scale if w is not None else None),
                market_value=v, is_swap=s)
        for n, t, w, _, v, s in raw
    ]
    return holdings, as_of


def aggregate_by_issuer(holdings: list[Holding]) -> list[Issuer]:
    """Combine swap and direct-equity lines per issuer; drop cash and collateral."""
    issuers: dict[str, Issuer] = {}
    for h in holdings:
        if NON_EQUITY_RE.search(h.name):
            continue
        key = issuer_key(h.name)
        if not key:
            continue
        iss = issuers.get(key)
        if iss is None:
            display = re.sub(r"\s+(total\s+return\s+)?swap.*$", "", h.name, flags=re.I).strip()
            iss = issuers[key] = Issuer(key=key, name=display, weight_pct=0.0,
                                        market_value=0.0, tickers=[], components=[])
        iss.weight_pct += h.weight_pct or 0.0
        iss.market_value += h.market_value or 0.0
        iss.components.append(("swap: " if h.is_swap else "equity: ") + h.name)
        if h.ticker and h.ticker not in iss.tickers:
            iss.tickers.append(h.ticker)
    ranked = sorted(issuers.values(), key=lambda i: (i.weight_pct, i.market_value), reverse=True)
    return [i for i in ranked if i.weight_pct > 0 or i.market_value > 0]


def fetch_msos_issuers(web: HttpClient, holdings_file: str | None, refresh: bool) -> tuple[list[Issuer], str, str | None]:
    if holdings_file:
        path = Path(holdings_file)
        data, filename = path.read_bytes(), path.name
        source = MSOS_HOLDINGS_URL  # the public file the local copy was downloaded from
        log.info("Using local holdings file %s", path.resolve())
    else:
        try:
            data = web.get(MSOS_HOLDINGS_URL, ttl=12 * 3600, use_cache=not refresh)
        except PipelineError as exc:
            raise FatalError(
                f"Could not download the MSOS holdings file ({exc}). Download it from {MSOS_FUND_PAGE} "
                "and pass --holdings-file PATH, or pass --tickers CURLF,TCNNF,GTBIF,..."
            ) from exc
        source, filename = MSOS_HOLDINGS_URL, "holdings.xlsx"
    holdings, as_of = parse_holdings(data, filename)
    issuers = aggregate_by_issuer(holdings)
    log.info("MSOS holdings: %d lines -> %d equity issuers (as of %s)", len(holdings), len(issuers), as_of or "n/a")
    return issuers, source, as_of


# --------------------------------------------------------------------------- #
# Stage 2 - SEC ticker / CIK resolution
# --------------------------------------------------------------------------- #


@dataclass
class SecEntity:
    cik: int
    ticker: str
    title: str


class SecDirectory:
    def __init__(self, rows: list[SecEntity]) -> None:
        self.by_ticker = {r.ticker.upper(): r for r in rows}
        self.by_key: dict[str, SecEntity] = {}
        self.by_cik: dict[int, SecEntity] = {}
        for r in rows:  # first ticker listed for a CIK is its primary ticker
            self.by_key.setdefault(issuer_key(r.title), r)
            self.by_cik.setdefault(r.cik, r)

    @classmethod
    def load(cls, sec: HttpClient, refresh: bool) -> SecDirectory:
        data = sec.get_json(SEC_TICKERS_URL, ttl=24 * 3600, use_cache=not refresh)
        rows = [SecEntity(int(v["cik_str"]), str(v["ticker"]), str(v["title"])) for v in data.values()]
        return cls(rows)

    def resolve(self, issuer: Issuer) -> SecEntity | None:
        for t in issuer.tickers:  # holdings tickers may be CSE ('CURA CN') or OTC ('CURLF')
            hit = self.by_ticker.get(t.split()[0].upper())
            if hit and issuer_key(hit.title).split()[:1] == issuer.key.split()[:1]:
                return hit
        if issuer.key in self.by_key:
            return self.by_key[issuer.key]
        first = issuer.key.split()[0]
        candidates = [k for k in self.by_key if k.split()[:1] == [first]]
        close = difflib.get_close_matches(issuer.key, candidates, n=1, cutoff=0.8)
        if close:
            return self.by_key[close[0]]
        if issuer.key in SEED_CIKS:
            cik = SEED_CIKS[issuer.key]
            known = self.by_cik.get(cik)
            return known or SecEntity(cik, "", issuer.name)
        return None


# --------------------------------------------------------------------------- #
# Stage 3 - Primary-source annual report retrieval
# --------------------------------------------------------------------------- #


@dataclass
class AnnualFiling:
    cik: int
    company: str
    ticker: str
    form: str
    accession: str
    filing_date: str
    report_date: str
    primary_document: str

    @property
    def folder(self) -> str:
        return f"https://www.sec.gov/Archives/edgar/data/{self.cik}/{self.accession.replace('-', '')}"

    @property
    def url(self) -> str:
        return f"{self.folder}/{self.primary_document}"

    @property
    def index_url(self) -> str:
        return f"{self.folder}/{self.accession}-index.htm"


def find_latest_annual_filing(sec: HttpClient, cik: int, refresh: bool) -> AnnualFiling | None:
    data = sec.get_json(SEC_SUBMISSIONS_URL.format(cik=cik), ttl=24 * 3600, use_cache=not refresh)
    recent = data.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    best: AnnualFiling | None = None
    for i, form in enumerate(forms):
        if form not in ANNUAL_FORMS:  # originals only; amendments are often Part III-only
            continue
        filing = AnnualFiling(
            cik=cik,
            company=data.get("name", ""),
            ticker=(data.get("tickers") or [""])[0],
            form=form,
            accession=recent["accessionNumber"][i],
            filing_date=recent["filingDate"][i],
            report_date=recent.get("reportDate", [""] * len(forms))[i] or "",
            primary_document=recent["primaryDocument"][i],
        )
        if best is None or filing.filing_date > best.filing_date:
            best = filing
    return best


# --------------------------------------------------------------------------- #
# Stage 4 - Footnote isolation
# --------------------------------------------------------------------------- #

def parse_html(html: bytes | str) -> BeautifulSoup:
    """Parse EDGAR HTML/XHTML (inline XBRL filings are XHTML) with lxml's lenient HTML parser."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", XMLParsedAsHTMLWarning)
        return BeautifulSoup(html, "lxml")


BLOCK_TAGS = {
    "p", "div", "br", "tr", "li", "ul", "ol", "table", "section", "article",
    "h1", "h2", "h3", "h4", "h5", "h6", "center", "blockquote", "pre", "hr",
}


def _is_hidden(el: Tag) -> bool:
    style = (el.get("style") or "").replace(" ", "").lower()
    return "display:none" in style or el.name in ("ix:header", "script", "style", "head")


def render_table(table: Tag) -> str:
    """Render an HTML table as pipe-delimited rows, gluing the '$', '(' and ')'
    fragments that EDGAR filers put in separate cells back onto their numbers."""
    lines = []
    for tr in table.find_all("tr"):
        cells = []
        for td in tr.find_all(["td", "th"]):
            t = " ".join(td.get_text(" ", strip=True).split())
            if t:
                cells.append(t)
        merged: list[str] = []
        for c in cells:
            if merged and re.fullmatch(r"\)?\s*%?\)?", c) and c.strip():
                merged[-1] += c.replace(" ", "")
            elif merged and merged[-1] in ("$", "(", "$(", "($"):
                merged[-1] += c
            else:
                merged.append(c)
        if merged:
            lines.append(" | ".join(merged))
    return "\n".join(lines)


def html_to_text(node: Tag) -> str:
    parts: list[str] = []

    def walk(el: Any) -> None:
        if isinstance(el, Comment):
            return
        if isinstance(el, NavigableString):
            parts.append(str(el))
            return
        if not isinstance(el, Tag) or _is_hidden(el):
            return
        if el.name == "table":
            parts.append("\n" + render_table(el) + "\n")
            return
        if el.name == "br":
            parts.append("\n")
            return
        block = el.name in BLOCK_TAGS
        if block:
            parts.append("\n")
        for child in el.children:
            walk(child)
        if block:
            parts.append("\n")

    walk(node)
    text = "".join(parts).replace("\xa0", " ").replace("​", "")
    lines = [" ".join(line.split()) for line in text.splitlines()]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()


def extract_ixbrl_text_block(soup: BeautifulSoup, concept_suffix: str) -> str | None:
    """Return the text of an Inline XBRL text block (following `continuedAt`
    chains), e.g. us-gaap:IncomeTaxDisclosureTextBlock - the issuer's own tag for
    the complete income-tax footnote."""
    pat = re.compile(r"(^|:)" + re.escape(concept_suffix) + r"$", re.I)
    el = soup.find("ix:nonnumeric", attrs={"name": pat})
    if el is None:
        return None
    chunks, seen = [html_to_text(el)], set()
    nxt = el.get("continuedat")
    while nxt and nxt not in seen:
        seen.add(nxt)
        cont = soup.find("ix:continuation", attrs={"id": nxt})
        if cont is None:
            break
        chunks.append(html_to_text(cont))
        nxt = cont.get("continuedat")
    text = "\n".join(c for c in chunks if c).strip()
    return text or None


TAX_HEADING_RE = re.compile(
    r"^(?:note\s+)?(?:\d{1,2}[a-z]?\s*[.:)\-–—]?\s*)?(?:provision\s+for\s+)?income\s+tax(?:es|ation)?$", re.I
)
SECTION_BREAK_RE = re.compile(
    r"^(item\s+\d+[a-z]?\b|(consolidated|combined)\s+(statements?|balance\s+sheets?)\b|"
    r"notes\s+to\s+(the\s+)?(consolidated|combined|financial))",
    re.I,
)
NOTE_HEADING_RE = re.compile(r"^(?:note\s+)?(\d{1,2}[a-z]?)\s*[.:)\-–—]\s*(.{3,90})$", re.I)


def _is_note_heading(line: str) -> bool:
    m = NOTE_HEADING_RE.match(line.strip())
    if not m or line.rstrip().endswith((".", ",", ";")):
        return False
    words = [w for w in re.findall(r"[A-Za-z]+", m.group(2)) if len(w) > 3]
    return bool(words) and all(w[0].isupper() for w in words)


def score_tax_note(text: str) -> int:
    low = text.lower()
    score = 0
    score += 4 if "280e" in low else 0
    score += 2 if "statutory" in low else 0
    score += 2 if "reconcil" in low else 0
    score += 1 if "effective tax rate" in low else 0
    score += 1 if "deferred tax" in low else 0
    score += 1 if "valuation allowance" in low else 0
    score += 1 if "unrecognized tax" in low or "uncertain tax" in low else 0
    if any("|" in ln and re.search(r"statutory|expected|computed|federal", ln, re.I)
           for ln in text.splitlines()):
        score += 5  # a rendered reconciliation table row
    return score


def extract_tax_note_heuristic(full_text: str, max_chars: int = 80_000) -> str | None:
    """Fallback for filings without Inline XBRL: find 'Income Taxes' headings and
    keep the segment (up to the next numbered note) that best looks like the
    footnote rather than the accounting-policy paragraph or MD&A."""
    lines = full_text.splitlines()
    best: tuple[int, int, str] | None = None
    for i, line in enumerate(lines):
        if not TAX_HEADING_RE.match(line.strip()):
            continue
        seg, size = [line], len(line)
        for nxt in lines[i + 1:]:
            stripped = nxt.strip()
            if TAX_HEADING_RE.match(stripped) or SECTION_BREAK_RE.match(stripped) or _is_note_heading(stripped):
                break
            seg.append(nxt)
            size += len(nxt) + 1
            if size > max_chars:
                break
        text = "\n".join(seg).strip()
        if len(text) < 300:
            continue
        cand = (score_tax_note(text), i, text)  # ties -> later occurrence (notes follow MD&A)
        if best is None or cand[:2] > best[:2]:
            best = cand
    if best is None or best[0] < 10:  # requires the reconciliation-table signal
        return None
    return best[2]


def _heading_context(table: Tag, max_chars: int = 600) -> str:
    bits: list[str] = []
    total = 0
    for s in table.find_all_previous(string=True):
        if isinstance(s, Comment):
            continue
        parent = s.parent
        if parent is not None and (parent.find_parent("table") is not None or parent.name == "table"):
            if bits:
                break
            continue
        if parent is not None and any(_is_hidden(p) for p in [parent, *parent.parents] if isinstance(p, Tag)):
            continue
        t = " ".join(str(s).split())
        if t:
            bits.append(t)
            total += len(t)
        if total > max_chars or len(bits) >= 15:
            break
    return " ".join(reversed(bits))


def extract_income_statement(soup: BeautifulSoup) -> str | None:
    """Locate the consolidated statement of operations (heading + table)."""
    best: tuple[int, int, str] | None = None
    for idx, table in enumerate(soup.find_all("table")):
        text = render_table(table)
        if len(text) < 150:
            continue
        low = text.lower()
        if not (re.search(r"net\s+(\(loss\)|loss|income)", low) and re.search(r"revenue|net sales", low)
                and re.search(r"income tax", low)):
            continue
        context = _heading_context(table)
        score = 1
        if re.search(r"statements?\s+of\s+(consolidated\s+)?(operations|income|loss|comprehensive)", context, re.I):
            score += 4
        score += 1 if "per share" in low else 0
        score += 1 if "gross profit" in low else 0
        cand = (score, -idx, f"{context}\n{text}")
        if best is None or cand[:2] > best[:2]:
            best = cand
    return best[2] if best else None


@dataclass
class FinancialSections:
    document_url: str
    tax_note: str
    tax_note_method: str
    income_statement: str | None


_SKIP_DOC_RE = re.compile(
    r"consent|cert|ex-?3[12]|ex-?2[134]|ex-?97|ex-?19|power|charter|code|press|earning|"
    r"release|\baif\b|xaif|annualinformation",
    re.I,
)


def rank_filing_documents(names: list[str], primary: str) -> list[str]:
    """Order a filing's HTML documents by how likely they are to contain the audited
    financial statements (10-K: the primary document; 40-F: an exhibit)."""
    def score(name: str) -> int:
        low = name.lower()
        s = 0
        if name == primary:
            s += 100
        if re.search(r"\d{8}(_d\d+)?\.html?$", low):
            s += 50  # inline XBRL instance naming, e.g. curlf-20251231.htm
        if re.search(r"financ|statements|audited|(^|[^a-z])fs([^a-z]|$)", low):
            s += 40
        if re.search(r"ex-?99", low):
            s += 10
        if _SKIP_DOC_RE.search(low):
            s -= 200
        return s

    htm = [n for n in names if n.lower().endswith((".htm", ".html")) and "index" not in n.lower()]
    return sorted(dict.fromkeys(htm), key=score, reverse=True)


def locate_financial_sections(sec: HttpClient, filing: AnnualFiling, max_docs: int = 6) -> FinancialSections:
    names = [filing.primary_document]
    try:
        idx = sec.get_json(f"{filing.folder}/index.json")
        names += [item["name"] for item in idx.get("directory", {}).get("item", [])]
    except PipelineError as exc:
        log.warning("Could not list documents for %s (%s); using primary document only", filing.accession, exc)
    candidates = [n for n in rank_filing_documents(names, filing.primary_document)][:max_docs]

    for doc in candidates:
        url = f"{filing.folder}/{doc}"
        try:
            html = sec.get(url)
        except PipelineError as exc:
            log.warning("  skip %s: %s", doc, exc)
            continue
        soup = parse_html(html)
        note = extract_ixbrl_text_block(soup, "IncomeTaxDisclosureTextBlock")
        method = "ixbrl:us-gaap:IncomeTaxDisclosureTextBlock"
        if not note or score_tax_note(note) < 3:
            note = extract_tax_note_heuristic(html_to_text(soup.body or soup))
            method = "heading-heuristic"
        if not note:
            log.debug("  no income-tax footnote in %s", doc)
            continue
        statement = extract_income_statement(soup)
        if statement is None:
            log.warning("  income statement not located in %s", doc)
        log.info("  footnote located in %s via %s (%s chars)", doc, method, f"{len(note):,}")
        return FinancialSections(url, note, method, statement)
    raise PipelineError(f"Income Taxes footnote not found in {filing.form} {filing.accession}")


# --------------------------------------------------------------------------- #
# Stage 5 - LLM extraction (strict JSON schema)
# --------------------------------------------------------------------------- #

LineCategory = Literal[
    "statutory", "section_280e", "nondeductible_other", "state_local",
    "valuation_allowance", "uncertain_tax_positions", "other",
]
ReserveLink = Literal["line_attributed", "reserve_attributed", "narrative_only", "none"]
Multiplier = Literal[1, 1000, 1000000]


class _Strict(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ReconciliationLine(_Strict):
    label: str
    category: LineCategory
    reserve_280e_link: ReserveLink
    amount: Optional[float]
    percent: Optional[float]
    increases_tax_expense: bool
    rationale: str


class IncomeStatementFigures(_Strict):
    unit_multiplier: Multiplier
    net_income_label: str
    net_income: Optional[float]
    pretax_income: Optional[float]
    income_tax_expense: Optional[float]
    total_revenue: Optional[float]


class RateReconciliation(_Strict):
    found: bool
    unit_multiplier: Multiplier
    presentation: Literal["amounts_and_percentages", "amounts_only", "percentages_only", "not_presented"]
    statutory_rate_percent: Optional[float]
    pretax_income: Optional[float]
    total_income_tax_expense: Optional[float]
    line_items: list[ReconciliationLine]


class TaxExtraction(_Strict):
    fiscal_year_end: str
    income_statement: IncomeStatementFigures
    rate_reconciliation: RateReconciliation
    section_280e_quotes: list[str]
    uncertain_tax_position_on_280e: bool
    extraction_notes: str


def _num(desc: str) -> dict[str, Any]:
    return {"anyOf": [{"type": "number"}, {"type": "null"}], "description": desc}


def _obj(props: dict[str, Any], desc: str = "") -> dict[str, Any]:
    out: dict[str, Any] = {"type": "object", "properties": props,
                           "required": list(props), "additionalProperties": False}
    if desc:
        out["description"] = desc
    return out


MULTIPLIER_SCHEMA = {"type": "integer", "enum": [1, 1000, 1000000],
                     "description": "Scale stated in the table header: 1000 for 'in thousands', 1000000 for 'in millions', else 1."}

EXTRACTION_SCHEMA: dict[str, Any] = _obj({
    "fiscal_year_end": {"type": "string", "description": "Most recent fiscal year end, YYYY-MM-DD."},
    "income_statement": _obj({
        "unit_multiplier": MULTIPLIER_SCHEMA,
        "net_income_label": {"type": "string", "description": "Verbatim label of the net income (loss) line used."},
        "net_income": _num("Consolidated net income (loss) as printed (loss negative), before allocation to non-controlling interests."),
        "pretax_income": _num("Income (loss) before income taxes as printed (loss negative)."),
        "income_tax_expense": _num("Provision for (benefit from) income taxes as printed; expense positive, benefit negative."),
        "total_revenue": _num("Total (net) revenue as printed."),
    }, "Figures from the consolidated statement of operations, most recent fiscal year only."),
    "rate_reconciliation": _obj({
        "found": {"type": "boolean"},
        "unit_multiplier": MULTIPLIER_SCHEMA,
        "presentation": {"type": "string", "enum": ["amounts_and_percentages", "amounts_only", "percentages_only", "not_presented"]},
        "statutory_rate_percent": _num("Federal statutory rate used, e.g. 21.0."),
        "pretax_income": _num("Pre-tax income (loss) the reconciliation starts from, as printed, if shown."),
        "total_income_tax_expense": _num("Total income tax expense (benefit) at the bottom of the reconciliation, as printed; expense positive."),
        "line_items": {
            "type": "array",
            "description": "Every reconciling line for the most recent year, in table order, excluding the final total line.",
            "items": _obj({
                "label": {"type": "string", "description": "Verbatim line label."},
                "category": {"type": "string", "enum": list(LineCategory.__args__)},  # type: ignore[attr-defined]
                "reserve_280e_link": {"type": "string", "enum": list(ReserveLink.__args__),  # type: ignore[attr-defined]
                                      "description": "For uncertain_tax_positions lines: how the footnote ties the "
                                                     "reserve to Section 280E. 'none' for every other line."},
                "amount": _num("Currency amount as printed, sign-normalised: positive increases income tax expense, negative decreases it. Null if not presented."),
                "percent": _num("Rate effect as printed, e.g. -45.3 for '(45.3)%'. Null if not presented."),
                "increases_tax_expense": {"type": "boolean"},
                "rationale": {"type": "string", "description": "One sentence justifying the category and reserve_280e_link; quote the text tying the line or reserve to 280E where applicable."},
            }),
        },
    }, "The statutory-to-effective tax rate reconciliation from the income taxes footnote."),
    "section_280e_quotes": {"type": "array", "items": {"type": "string"},
                            "description": "Up to three verbatim sentences from the footnote that discuss Section 280E."},
    "uncertain_tax_position_on_280e": {"type": "boolean"},
    "extraction_notes": {"type": "string"},
})

SYSTEM_PROMPT = """\
You are a forensic tax accountant supporting private-equity due diligence on U.S. \
multi-state cannabis operators. You transcribe figures from audited SEC filings \
exactly as printed. You never estimate, interpolate, net, or sum figures that the \
filing does not state; when a value is not presented you return null and explain \
in extraction_notes. Your output feeds a deterministic model, so faithful \
transcription and correct line classification matter more than anything else."""

INSTRUCTIONS = """\
Work only with the fiscal year ended {period} (ignore comparative-year columns). \
Tables are rendered as pipe-delimited rows; parentheses denote negative numbers \
and "—" denotes zero. Report numbers as printed (do not rescale them); express \
scale through unit_multiplier.

1. income_statement - from the consolidated statement of operations:
   - net_income: the consolidated net income (loss) line before allocation to \
non-controlling interests ("Net loss", "Net (loss) income"). Use the line \
attributable to the company only if no consolidated line exists, and say so in \
extraction_notes.
   - pretax_income, income_tax_expense (expense positive, benefit negative), \
total_revenue.

2. rate_reconciliation - the statutory-to-effective tax rate reconciliation in \
the income taxes footnote:
   - line_items: every reconciling line in table order, starting with tax at the \
statutory rate and excluding the final total (report the total in \
total_income_tax_expense). Where a category is disaggregated into sub-lines (as \
ASU 2023-09 requires, e.g. "Nontaxable or nondeductible items" -> "Section 280E", \
"Share-based compensation"), report the sub-lines instead of the subtotal so \
nothing is counted twice.
   - amount: sign-normalised so a positive number increases income tax expense (or \
reduces a benefit). Null if the table shows percentages only.
   - category:
       "section_280e" - the label, a parenthetical, a footnote marker on the line, or \
the narrative explicitly attributes this specific line to IRC Section 280E \
(e.g. "Section 280E", "Nondeductible expenses - 280E", "Non-deductible expenses \
(primarily IRC 280E)"). Not for uncertain-tax-position reserve lines; see \
reserve_280e_link.
       "nondeductible_other" - nondeductible or permanent items not explicitly tied \
to 280E (share-based compensation, goodwill impairment, fair-value changes, ...).
       "statutory", "state_local", "valuation_allowance", "uncertain_tax_positions", \
"other" - as named.
   - reserve_280e_link: many operators file returns as if 280E does not apply and \
record the disputed tax as an uncertain-tax-position (UTP / unrecognized tax \
benefit) reserve, so the 280E cost appears in the reconciliation as a change in \
that reserve. Classify a reserve line as "uncertain_tax_positions" even when it \
is attributed to 280E, and record the tie here:
       "line_attributed" - a label, parenthetical or footnote marker on this line, \
or the narrative, attributes this line itself wholly or primarily to Section 280E \
(e.g. a marker reading "Primarily related to the Company's Section 280E Position").
       "reserve_attributed" - the footnote identifies the reserve this line changes, \
its rollforward, or a stated dollar portion of it as the company's 280E position \
(e.g. "recorded an uncertain tax liability for positions that challenge its \
liability under Section 280E ... reflected in the tables below", "$X of the \
liability relates to the 280E position", a rollforward row labelled 280E).
       "narrative_only" - the footnote says only in general terms that 280E gives \
rise to unrecognized tax benefits, without identifying the reserve, its \
rollforward or any amount as the 280E position.
       "none" - the footnote does not tie the reserve to 280E.
     Use "none" for every line that is not an uncertain_tax_positions line.
   - If there is no reconciliation table, set found=false, \
presentation="not_presented" and leave line_items empty.

3. section_280e_quotes: up to three verbatim sentences discussing Section 280E.

4. uncertain_tax_position_on_280e: true if the footnote says the company has filed, \
or intends to file, returns or refund claims treating 280E as inapplicable and \
carries an uncertain-tax-position liability for it.

5. extraction_notes: anything an analyst must know (percent-only presentation, \
280E embedded in a combined line, restatements, figures not presented)."""


class LLMExtractor:
    """Calls Claude with a strict JSON schema, validates the result, caches it on
    disk, and separates retryable failures from fatal configuration errors."""

    def __init__(self, *, model: str, effort: str, cache_dir: Path | None,
                 client: Any = None, max_attempts: int = 4, sleep=time.sleep) -> None:
        if client is None:
            if anthropic is None:
                raise FatalError("The 'anthropic' package is not installed: pip install -r requirements.txt")
            api_key = os.environ.get("MSOS_ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")
            if not (api_key or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
                raise FatalError(
                    "MSOS_ANTHROPIC_API_KEY is not set. Export it in your shell or add it as an "
                    "environment variable in your runtime; never paste it into code or chat."
                )
            # The SDK already retries 408/409/429/5xx and connection errors with
            # exponential backoff; max_retries widens that envelope.
            kwargs: dict[str, Any] = {"api_key": api_key} if api_key else {}
            client = anthropic.Anthropic(max_retries=5, timeout=900.0, **kwargs)
        self.client = client
        self.model = model
        self.effort = effort
        self.cache_dir = cache_dir
        self.max_attempts = max_attempts
        self._sleep = sleep
        if cache_dir:
            cache_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def build_prompt(filing: AnnualFiling, company: str, sections: FinancialSections) -> str:
        period = filing.report_date or "the most recent fiscal year"
        statement = sections.income_statement or "[Statement of operations not located - return nulls for income_statement figures.]"
        return (
            f"<filing_metadata>\nCompany: {company}\nTicker: {filing.ticker}\nCIK: {filing.cik}\n"
            f"Form: {filing.form}\nAccession: {filing.accession}\nPeriod of report: {period}\n"
            f"Source: {sections.document_url}\n</filing_metadata>\n\n"
            f"<consolidated_statement_of_operations>\n{statement}\n</consolidated_statement_of_operations>\n\n"
            f"<income_taxes_footnote>\n{sections.tax_note}\n</income_taxes_footnote>\n\n"
            + INSTRUCTIONS.format(period=period)
        )

    def _cache_key(self, prompt: str) -> str:
        blob = json.dumps({"v": PROMPT_VERSION, "m": self.model, "e": self.effort,
                           "s": SYSTEM_PROMPT, "schema": EXTRACTION_SCHEMA, "p": prompt}, sort_keys=True)
        return hashlib.sha256(blob.encode()).hexdigest()[:40]

    def _call(self, prompt: str) -> str:
        with self.client.beta.messages.stream(
            model=self.model,
            max_tokens=32000,
            # Server-side refusal fallback: if a safety classifier declines, the API
            # re-runs the request on Anthropic's recommended fallback model.
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            thinking={"type": "adaptive"},
            output_config={"effort": self.effort,
                           "format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": prompt}],
        ) as stream:
            message = stream.get_final_message()

        if message.stop_reason == "refusal":
            raise PipelineError(f"Model declined the request: {getattr(message, 'stop_details', None)}")
        if message.stop_reason == "max_tokens":
            raise PipelineError("Model output truncated at max_tokens")
        text_parts: list[str] = []
        for block in message.content:
            if block.type == "fallback":  # discard any text produced before a fallback switch
                text_parts = []
            elif block.type == "text":
                text_parts.append(block.text)
        if not text_parts:
            raise PipelineError(f"No text content in response (stop_reason={message.stop_reason})")
        log.debug("LLM request %s usage=%s", getattr(message, "_request_id", "?"), getattr(message, "usage", None))
        return "".join(text_parts)

    def extract(self, filing: AnnualFiling, company: str, sections: FinancialSections) -> TaxExtraction:
        prompt = self.build_prompt(filing, company, sections)
        key = self._cache_key(prompt)
        cache_file = self.cache_dir / f"{key}.json" if self.cache_dir else None
        if cache_file and cache_file.exists():
            log.info("  LLM extraction cache hit")
            return TaxExtraction.model_validate_json(cache_file.read_text())

        last: Exception | None = None
        for attempt in range(1, self.max_attempts + 1):
            try:
                raw = self._call(prompt)
                result = TaxExtraction.model_validate_json(raw)
                if cache_file:
                    cache_file.write_text(result.model_dump_json(indent=2))
                return result
            except ValidationError as exc:  # should not happen with structured outputs
                last, delay = exc, 2.0
            except PipelineError as exc:
                last, delay = exc, 5.0
            except Exception as exc:  # classify SDK errors (most-specific first)
                if anthropic is None or not isinstance(exc, anthropic.AnthropicError):
                    raise
                if isinstance(exc, (anthropic.AuthenticationError, anthropic.PermissionDeniedError)):
                    raise FatalError(f"Anthropic API rejected the credentials: {exc}") from exc
                if isinstance(exc, anthropic.NotFoundError):
                    raise FatalError(f"Unknown model or endpoint '{self.model}': {exc}") from exc
                if isinstance(exc, anthropic.RequestTooLargeError):
                    raise PipelineError(f"Footnote exceeds request size limit: {exc}") from exc
                if isinstance(exc, anthropic.BadRequestError):
                    raise FatalError(f"Anthropic API rejected the request: {exc}") from exc
                if isinstance(exc, anthropic.RateLimitError):
                    retry_after = _parse_retry_after(exc.response.headers.get("retry-after"))
                    last, delay = exc, max(retry_after or 0.0, 30.0 * attempt)
                elif isinstance(exc, anthropic.APIStatusError) and exc.status_code >= 500:
                    last, delay = exc, 20.0 * attempt
                elif isinstance(exc, anthropic.APIConnectionError):  # includes timeouts
                    last, delay = exc, 15.0 * attempt
                else:
                    raise PipelineError(f"Anthropic API error: {exc}") from exc
            if attempt < self.max_attempts:
                log.warning("  LLM attempt %d/%d failed (%s); retrying in %.0fs",
                            attempt, self.max_attempts, type(last).__name__, delay)
                self._sleep(delay)
        raise PipelineError(f"LLM extraction failed after {self.max_attempts} attempts: {last}")


# --------------------------------------------------------------------------- #
# Stage 6 - Financial logic
# --------------------------------------------------------------------------- #


@dataclass
class PenaltySelection:
    amount_usd: float | None
    method: str
    lines: list[str]
    flags: list[str]
    evidence: list[str] = field(default_factory=list)


# Reserve lines that only carry interest and penalties are not the 280E tax itself.
_INTEREST_PENALTY_RE = re.compile(r"interest|penalt", re.I)
_RESERVE_LABEL_RE = re.compile(r"uncertain|unrecogni[sz]ed|reserve|\bUT[BP]s?\b", re.I)


def _line_usd(line: ReconciliationLine, multiplier: int, pretax_usd: float | None) -> tuple[float | None, str | None]:
    if line.amount is not None:
        return line.amount * multiplier, None
    if line.percent is not None and pretax_usd is not None:
        magnitude = abs(line.percent / 100.0 * pretax_usd)
        return (magnitude if line.increases_tax_expense else -magnitude), "PERCENT_DERIVED"
    return None, None


def select_280e_penalty(x: TaxExtraction, *, narrative_reserves: bool = False) -> PenaltySelection:
    """Deterministic selection of the 280E penalty from the extracted reconciliation.

    Preference order:
      1. lines the filing explicitly attributes to Section 280E;
      2. otherwise, uncertain-tax-position reserve lines, for operators that file as
         if 280E does not apply and reserve for the disputed tax, when the footnote
         attributes the line, or the reserve it changes, to 280E (or, with
         `narrative_reserves`, merely says 280E gives rise to it). Lines that carry
         only interest and penalties are excluded;
      3. otherwise, if the company reserves for 280E, nothing: the 280E cost sits in
         a reserve the filing does not quantify, so nondeductible lines would miss it;
      4. otherwise, generic nondeductible-expense lines, but only when the footnote
         narrative discusses 280E (flagged as a proxy - may include non-280E items);
      5. otherwise, nothing: the penalty is reported as not disclosed.
    """
    rr, inc = x.rate_reconciliation, x.income_statement
    flags: list[str] = []
    if not rr.found or not rr.line_items:
        return PenaltySelection(None, "no_reconciliation", [], ["NO_RATE_RECONCILIATION"])

    pretax_usd = None
    if rr.pretax_income is not None:
        pretax_usd = rr.pretax_income * rr.unit_multiplier
    elif inc.pretax_income is not None:
        pretax_usd = inc.pretax_income * inc.unit_multiplier

    reserves = [ln for ln in rr.line_items if ln.category == "uncertain_tax_positions"
                and not (_INTEREST_PENALTY_RE.search(ln.label) and not _RESERVE_LABEL_RE.search(ln.label))]
    accepted = {"line_attributed", "reserve_attributed"} | ({"narrative_only"} if narrative_reserves else set())
    explicit = [ln for ln in rr.line_items if ln.category == "section_280e"]
    tied = [ln for ln in reserves if ln.reserve_280e_link in accepted]
    if explicit:
        chosen, method = explicit, "explicit_280e"
    elif tied:
        chosen, method = tied, "reserve_280e"
        flags.append("280E_FROM_RESERVE_LINE")
        if any(ln.reserve_280e_link == "narrative_only" for ln in tied):
            flags.append("RESERVE_NARRATIVE_TIE_ONLY")
        if any(_INTEREST_PENALTY_RE.search(ln.label) for ln in tied):
            flags.append("RESERVE_INCLUDES_INTEREST_PENALTIES")
    elif any(ln.reserve_280e_link != "none" for ln in reserves) or (reserves and x.uncertain_tax_position_on_280e):
        return PenaltySelection(None, "reserve_not_attributed", [], ["280E_IN_RESERVE_NOT_QUANTIFIED"],
                                [ln.rationale for ln in reserves])
    else:
        proxy = [ln for ln in rr.line_items if ln.category == "nondeductible_other"]
        if proxy and x.section_280e_quotes:
            chosen, method = proxy, "nondeductible_proxy"
            flags.append("280E_NOT_SEPARATELY_LABELLED")
        else:
            return PenaltySelection(None, "not_disclosed", [], ["NO_280E_LINE"])

    total, labels, evidence = 0.0, [], []
    for ln in chosen:
        usd, flag = _line_usd(ln, rr.unit_multiplier, pretax_usd)
        if usd is None:
            flags.append(f"UNQUANTIFIED_LINE:{ln.label}")
            continue
        if flag and flag not in flags:
            flags.append(flag)
        total += usd
        labels.append(ln.label)
        evidence.append(ln.rationale)
    if not labels:
        return PenaltySelection(None, method, [], flags + ["NO_QUANTIFIABLE_280E_LINE"])
    if total < 0 and method == "reserve_280e":  # a net release of the reserve is not a 280E charge
        return PenaltySelection(None, method, labels, flags + ["RESERVE_NET_RELEASE"], evidence)
    if total < 0:  # 280E disallows deductions; it can only increase tax expense
        flags.append("SIGN_NORMALISED")
        total = abs(total)
    return PenaltySelection(total, method, labels, flags, evidence)


def reconciliation_ties(x: TaxExtraction) -> bool | None:
    rr = x.rate_reconciliation
    if not rr.found or rr.total_income_tax_expense is None or not rr.line_items:
        return None
    if any(ln.amount is None for ln in rr.line_items):
        return None
    diff = abs(sum(ln.amount for ln in rr.line_items) - rr.total_income_tax_expense)  # type: ignore[misc]
    return diff <= max(0.01 * abs(rr.total_income_tax_expense), 2.0)  # rounding in printed units


def xbrl_net_income(sec: HttpClient, filing: AnnualFiling, period_end: str) -> dict[str, float]:
    """Net income facts the issuer tagged in this filing (non-dimensional, ~12-month duration)."""
    data = sec.get_json(SEC_COMPANYFACTS_URL.format(cik=filing.cik), ttl=24 * 3600)
    out: dict[str, float] = {}
    for tag in ("ProfitLoss", "NetIncomeLoss"):
        for fact in data.get("facts", {}).get("us-gaap", {}).get(tag, {}).get("units", {}).get("USD", []):
            if fact.get("accn") != filing.accession or fact.get("end") != period_end or not fact.get("start"):
                continue
            days = (date.fromisoformat(fact["end"]) - date.fromisoformat(fact["start"])).days
            if 350 <= days <= 380:
                out[tag] = float(fact["val"])
                break
    return out


@dataclass
class CompanyResult:
    rank: int
    issuer: str
    msos_weight_pct: float
    company: str = ""
    ticker: str = ""
    cik: int | None = None
    form: str = ""
    accession: str = ""
    fiscal_year_end: str = ""
    filing_url: str = ""
    document_url: str = ""
    footnote_method: str = ""
    reported_net_income: float | None = None
    net_income_label: str = ""
    penalty_280e: float | None = None
    penalty_method: str = ""
    penalty_lines: list[str] = field(default_factory=list)
    penalty_evidence: list[str] = field(default_factory=list)
    pro_forma_net_income: float | None = None
    pretax_income: float | None = None
    income_tax_expense: float | None = None
    total_revenue: float | None = None
    effective_tax_rate: float | None = None
    pro_forma_effective_tax_rate: float | None = None
    reconciliation_ties: bool | None = None
    xbrl_check: str = "not run"
    uncertain_tax_position_on_280e: bool | None = None
    section_280e_quotes: list[str] = field(default_factory=list)
    extraction_notes: str = ""
    flags: list[str] = field(default_factory=list)
    status: str = "pending"
    error: str = ""


def apply_extraction(res: CompanyResult, x: TaxExtraction, *, narrative_reserves: bool = False) -> None:
    inc, rr = x.income_statement, x.rate_reconciliation
    m = inc.unit_multiplier
    res.fiscal_year_end = x.fiscal_year_end or res.fiscal_year_end
    res.net_income_label = inc.net_income_label
    res.reported_net_income = inc.net_income * m if inc.net_income is not None else None
    res.pretax_income = inc.pretax_income * m if inc.pretax_income is not None else None
    res.income_tax_expense = inc.income_tax_expense * m if inc.income_tax_expense is not None else None
    res.total_revenue = inc.total_revenue * m if inc.total_revenue is not None else None
    if res.income_tax_expense is None and rr.total_income_tax_expense is not None:
        res.income_tax_expense = rr.total_income_tax_expense * rr.unit_multiplier

    sel = select_280e_penalty(x, narrative_reserves=narrative_reserves)
    res.penalty_280e, res.penalty_method, res.penalty_lines = sel.amount_usd, sel.method, sel.lines
    res.penalty_evidence = sel.evidence
    res.flags.extend(sel.flags)
    res.reconciliation_ties = reconciliation_ties(x)
    if res.reconciliation_ties is False:
        res.flags.append("RECONCILIATION_DOES_NOT_TIE")
    if (rr.total_income_tax_expense is not None and res.income_tax_expense is not None
            and inc.income_tax_expense is not None):
        recon_total = rr.total_income_tax_expense * rr.unit_multiplier
        if abs(recon_total - res.income_tax_expense) > max(0.01 * abs(res.income_tax_expense), 2.0 * m):
            res.flags.append("RECON_TOTAL_NE_INCOME_STATEMENT_TAX")
    res.uncertain_tax_position_on_280e = x.uncertain_tax_position_on_280e
    res.section_280e_quotes = x.section_280e_quotes
    res.extraction_notes = x.extraction_notes

    if res.reported_net_income is None:
        raise PipelineError("Net income not extracted")
    # Pro Forma Net Income = Reported Net Income + Section 280E penalty.
    res.pro_forma_net_income = res.reported_net_income + (res.penalty_280e or 0.0)
    if res.pretax_income:
        if res.income_tax_expense is not None:
            res.effective_tax_rate = res.income_tax_expense / res.pretax_income
            res.pro_forma_effective_tax_rate = (res.income_tax_expense - (res.penalty_280e or 0.0)) / res.pretax_income


def apply_xbrl_check(res: CompanyResult, facts: dict[str, float]) -> None:
    if not facts:
        res.xbrl_check = "no XBRL net income fact for this filing"
        return
    for tag, val in facts.items():
        if res.reported_net_income is not None and abs(val - res.reported_net_income) <= max(0.005 * abs(val), 1_000):
            res.xbrl_check = f"ties to us-gaap:{tag}"
            return
    res.xbrl_check = "MISMATCH vs " + ", ".join(f"us-gaap:{t}={fmt_usd(v)}" for t, v in facts.items())
    res.flags.append("NET_INCOME_XBRL_MISMATCH")


# --------------------------------------------------------------------------- #
# Stage 7 - Reporting
# --------------------------------------------------------------------------- #


def fmt_usd(v: float | None) -> str:
    if v is None:
        return "n/a"
    a = abs(v)
    if a >= 1e9:
        s = f"${a / 1e9:,.2f}B"
    elif a >= 1e6:
        s = f"${a / 1e6:,.1f}M"
    elif a >= 1e3:
        s = f"${a / 1e3:,.1f}K"
    else:
        s = f"${a:,.0f}"
    return f"({s})" if v < 0 else s


def _is_link(s: str) -> bool:
    return s.startswith(("http://", "https://", "file://"))


def fmt_pct(v: float | None) -> str:
    return "n/a" if v is None else f"{v * 100:.1f}%"


_NUMBER_WORDS = {1: "one", 2: "two", 3: "three", 4: "four", 5: "five", 6: "six", 7: "seven",
                 8: "eight", 9: "nine", 10: "ten"}


def _join(items: list[str]) -> str:
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " and " + items[-1]


def build_executive_summary(results: list[CompanyResult], holdings_source: str, as_of: str | None) -> str:
    ok = [r for r in results if r.status == "ok" and r.reported_net_income is not None]
    if not ok:
        return "_No company completed the pipeline; see errors above._"
    n = len(ok)
    names = _join([f"{r.company or r.issuer} ({r.ticker})" for r in ok])
    n_word = _NUMBER_WORDS.get(n, str(n))
    weight = sum(r.msos_weight_pct for r in ok)
    fys = Counter(r.fiscal_year_end[:4] for r in ok if r.fiscal_year_end)
    fy = f"FY{fys.most_common(1)[0][0]} " if fys else ""
    ni = sum(r.reported_net_income for r in ok)  # type: ignore[misc]
    pen = sum(r.penalty_280e or 0.0 for r in ok)
    pf = ni + pen
    forms = sorted({r.form for r in ok})
    before = sum(1 for r in ok if (r.reported_net_income or 0) > 0)
    after = sum(1 for r in ok if (r.pro_forma_net_income or 0) > 0)

    rev_clause = ""
    if all(r.total_revenue for r in ok):
        rev = sum(r.total_revenue for r in ok)  # type: ignore[misc]
        rev_clause = f", equivalent to {pen / rev:.1%} of their combined {fmt_usd(rev)} of revenue"
    if ni < 0 <= pf:
        outcome = (f"the cohort swings from an aggregate net loss of {fmt_usd(-ni)} to a pro forma "
                   f"net profit of {fmt_usd(pf)}")
    elif ni < 0 and pf < 0:
        outcome = (f"on a pro forma basis the cohort's aggregate net loss narrows from {fmt_usd(-ni)} to "
                   f"{fmt_usd(-pf)}, a {pen / abs(ni):.0%} reduction")
    else:
        outcome = (f"on a pro forma basis aggregate net income rises from {fmt_usd(ni)} to {fmt_usd(pf)}"
                   + (f", a {pen / ni:.0%} uplift" if ni > 0 else ""))
    if after > before:
        breadth = f"with the number of profitable operators rising from {before} to {after} of {n}"
    else:
        breadth = f"with {after} of {n} operators profitable pro forma, unchanged from reported"
    notes = []
    reserve = [r for r in ok if r.penalty_method == "reserve_280e" and r.penalty_280e]
    if reserve:
        one = len(reserve) == 1
        notes.append(f"for {_join([r.company or r.issuer for r in reserve])}, which {'files' if one else 'file'} "
                     "as if 280E does not apply, the figure is the change in the uncertain-tax-position reserve "
                     f"{'its footnote attributes' if one else 'their footnotes attribute'} to 280E, which can "
                     "include interest, penalties and prior-year positions")
        narrative = [r for r in reserve if "RESERVE_NARRATIVE_TIE_ONLY" in r.flags]
        if narrative:
            notes.append(f"{_join([r.company or r.issuer for r in narrative])} "
                         f"{'ties its reserve' if len(narrative) == 1 else 'tie their reserves'} to 280E only "
                         "in general terms")
    unquantified = [r for r in ok if r.penalty_method == "reserve_not_attributed"]
    if unquantified:
        one = len(unquantified) == 1
        notes.append(f"{_join([r.company or r.issuer for r in unquantified])} "
                     f"{'reserves' if one else 'reserve'} for 280E without attributing an amount to it and "
                     f"{'is' if one else 'are'} carried at zero")
    proxy = [r for r in ok if r.penalty_method == "nondeductible_proxy"]
    if proxy:
        notes.append(f"{_join([r.company or r.issuer for r in proxy])} "
                     f"{'discloses' if len(proxy) == 1 else 'disclose'} 280E only within a broader "
                     "nondeductible-expense line, flagged in the appendix")
    notes_clause = f" ({'; '.join(notes)})" if notes else ""

    paragraph = (
        "Because standard financial-data APIs report only total income tax expense and never isolate the "
        "Section 280E charge, I engineered a proprietary Python pipeline that programmatically pulls the "
        "AdvisorShares Pure US Cannabis ETF (MSOS) holdings file, aggregates total-return-swap and direct-equity "
        "exposure by issuer, resolves each operator to its SEC CIK, retrieves its most recent audited annual "
        f"report ({'/'.join(forms)}) from EDGAR, isolates the Income Taxes footnote from the Inline XBRL, and uses a "
        "schema-constrained LLM to transcribe the statutory-to-effective tax rate reconciliation line by line, "
        "while deterministic code selects the 280E line, verifies that the reconciliation ties to reported tax "
        "expense and cross-checks net income against the issuer's XBRL facts. "
        f"Across the {n_word} largest SEC-reporting operators in the fund ({names}"
        + (f"; together {weight:.1f}% of MSOS net assets, swap and direct exposure combined" if weight > 0 else "")
        + f"), the {fy}footnotes disclose an aggregate Section 280E penalty of "
        f"{fmt_usd(pen)}{rev_clause}{notes_clause}. "
        "Section 280E denies every deduction other than cost of goods sold to a business trafficking in a "
        "Schedule I or II substance, so these operators are taxed on gross profit rather than operating income; "
        "rescheduling to Schedule III ends that disallowance and restores the deductibility of SG&A. Because the "
        "280E charge in the reconciliation is the tax on those disallowed deductions, eliminating it reduces "
        "income tax expense dollar for dollar with no change to revenue, gross margin or operating costs, so the "
        f"penalty drops straight to the bottom line: {outcome}, a {fmt_usd(pen)} improvement in "
        f"aggregate profitability, {breadth}."
    )
    holdings_label = f"MSOS holdings file{f' (as of {as_of})' if as_of else ''}"
    sources = ([f"[{holdings_label}]({holdings_source})"] if _is_link(holdings_source) else []) + [
        f"[{r.company or r.issuer} {r.form} FY{r.fiscal_year_end[:4]}]({r.document_url or r.filing_url})" for r in ok
    ]
    return paragraph + "\n\nSources: " + "; ".join(sources) + "."


def render_markdown(results: list[CompanyResult], issuers: list[Issuer], summary: str,
                    holdings_source: str, as_of: str | None, model: str) -> str:
    ok = [r for r in results if r.status == "ok"]
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    out = [
        "# MSOS Section 280E Diligence - Pro Forma Schedule III Net Income",
        "",
        f"_Generated {ts} by `msos_280e_diligence.py` · extraction model `{model}` · "
        "holdings source: " + (f"[{holdings_source}]({holdings_source})" if _is_link(holdings_source) else holdings_source)
        + (f" (as of {as_of})" if as_of else "") + "_",
        "",
        "## Executive summary",
        "",
        summary,
        "",
        "## Results",
        "",
        "| # | Company | Ticker | Form | FY end | Reported net income | 280E penalty | Pro forma net income | ETR reported → pro forma | Checks |",
        "|---|---|---|---|---|---:|---:|---:|---|---|",
    ]
    for r in results:
        if r.status != "ok":
            out.append(f"| {r.rank} | {r.company or r.issuer} | {r.ticker} | {r.form} | | | | | | ERROR: {r.error} |")
            continue
        checks = [r.xbrl_check]
        if r.reconciliation_ties is not None:
            checks.append("reconciliation ties" if r.reconciliation_ties else "reconciliation does NOT tie")
        checks += r.flags
        out.append(
            f"| {r.rank} | [{r.company}]({r.document_url or r.filing_url}) | {r.ticker} | {r.form} | {r.fiscal_year_end} | "
            f"{fmt_usd(r.reported_net_income)} | {fmt_usd(r.penalty_280e)} | {fmt_usd(r.pro_forma_net_income)} | "
            f"{fmt_pct(r.effective_tax_rate)} → {fmt_pct(r.pro_forma_effective_tax_rate)} | {'; '.join(checks)} |"
        )
    if ok:
        ni = sum(r.reported_net_income or 0 for r in ok)
        pen = sum(r.penalty_280e or 0 for r in ok)
        out.append(f"| | **Aggregate** | | | | **{fmt_usd(ni)}** | **{fmt_usd(pen)}** | **{fmt_usd(ni + pen)}** | | |")

    out += ["", "## 280E line items used", ""]
    for r in ok:
        lines = "; ".join(f"“{ln}”" for ln in r.penalty_lines) or "none"
        out.append(f"- **{r.company}** ({r.penalty_method}): {lines}. Footnote located via `{r.footnote_method}`."
                   + (" Company carries an uncertain tax position on 280E." if r.uncertain_tax_position_on_280e else ""))
        for q in r.section_280e_quotes[:2]:
            out.append(f"  - > {q}")
        for ev in r.penalty_evidence:
            out.append(f"  - Line evidence: {ev}")
        if r.extraction_notes:
            out.append(f"  - Extraction notes: {r.extraction_notes}")

    out += ["", "## Target selection (MSOS issuers, swap + equity exposure combined)", "",
            "| MSOS rank | Issuer | Weight | SEC ticker | CIK | Status |", "|---|---|---:|---|---|---|"]
    for i, iss in enumerate(issuers, 1):
        if iss.status == "pending":
            continue
        out.append(f"| {i} | {iss.name} | {iss.weight_pct:.2f}% | {iss.sec_ticker or ''} | {iss.cik or ''} | {iss.status} |")

    out += [
        "", "## Methodology and limitations", "",
        "- **Pro forma definition.** Pro Forma Net Income = Reported Net Income + the tax charge the rate "
        "reconciliation attributes to Section 280E: an explicit 280E line where one exists, otherwise the "
        "change in the uncertain-tax-position reserve the footnote attributes to 280E (operators that file as "
        "if 280E does not apply book the disputed tax there). It removes that charge only; it does not model "
        "state conformity, deferred-tax remeasurement, release of reserves accrued in prior years, or "
        "second-order effects on pricing and competition. A reserve line can include interest, penalties and "
        "positions for prior years, so it can differ from the current year's 280E cost (flags "
        "`280E_FROM_RESERVE_LINE`, `RESERVE_INCLUDES_INTEREST_PENALTIES`).",
        "- **Line selection is deterministic.** The LLM only transcribes and classifies reconciliation lines and "
        "states how the footnote ties each reserve to 280E; code selects, in order: explicit 280E lines; reserve "
        "lines the footnote attributes to 280E, directly or as the company's 280E position (lines holding only interest "
        "and penalties are excluded; reserves tied to 280E only in narrative count only with "
        "`--narrative-reserves`, flag `RESERVE_NARRATIVE_TIE_ONLY`); nothing, when the company reserves for "
        "280E without attributing an amount to it (flag `280E_IN_RESERVE_NOT_QUANTIFIED`, penalty carried at "
        "zero); and, for companies with no 280E reserve, all nondeductible lines as a proxy when the footnote "
        "discusses 280E (flag `280E_NOT_SEPARATELY_LABELLED`), which can include unrelated permanent differences.",
        "- **Controls.** Reconciliation lines must sum to total tax expense; net income is cross-checked "
        "against the issuer's XBRL `ProfitLoss`/`NetIncomeLoss` facts for the same accession.",
        "- **Filer universe.** Curaleaf, Cresco Labs and Glass House file Form 40-F (U.S. GAAP financial "
        "statements) rather than 10-K; the pipeline treats both as the audited annual report.",
    ]
    return "\n".join(out) + "\n"


def write_outputs(out_dir: Path, results: list[CompanyResult], issuers: list[Issuer], summary: str,
                  holdings_source: str, as_of: str | None, model: str) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "results.json").write_text(json.dumps([asdict(r) for r in results], indent=2))
    cols = ["rank", "company", "ticker", "cik", "form", "accession", "fiscal_year_end", "msos_weight_pct",
            "reported_net_income", "penalty_280e", "pro_forma_net_income", "penalty_method", "pretax_income",
            "income_tax_expense", "total_revenue", "effective_tax_rate", "pro_forma_effective_tax_rate",
            "reconciliation_ties", "xbrl_check", "flags", "status", "error", "filing_url"]
    with (out_dir / "results.csv").open("w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        for r in results:
            row = asdict(r)
            w.writerow(["; ".join(row[c]) if isinstance(row[c], list) else row[c] for c in cols])
    (out_dir / "executive_summary.md").write_text(summary + "\n")
    (out_dir / "report.md").write_text(render_markdown(results, issuers, summary, holdings_source, as_of, model))


def print_table(results: list[CompanyResult]) -> None:
    header = f"{'#':>2}  {'Company':<30} {'Ticker':<7} {'Form':<5} {'Reported NI':>14} {'280E penalty':>14} {'Pro forma NI':>14}"
    print("\n" + header + "\n" + "-" * len(header))
    for r in results:
        if r.status != "ok":
            print(f"{r.rank:>2}  {(r.company or r.issuer)[:30]:<30} {r.ticker:<7} {r.form:<5} ERROR: {r.error[:60]}")
            continue
        print(f"{r.rank:>2}  {r.company[:30]:<30} {r.ticker:<7} {r.form:<5} "
              f"{fmt_usd(r.reported_net_income):>14} {fmt_usd(r.penalty_280e):>14} {fmt_usd(r.pro_forma_net_income):>14}")
    ok = [r for r in results if r.status == "ok"]
    if ok:
        ni = sum(r.reported_net_income or 0 for r in ok)
        pen = sum(r.penalty_280e or 0 for r in ok)
        print("-" * len(header))
        print(f"{'':>2}  {'AGGREGATE':<30} {'':<7} {'':<5} {fmt_usd(ni):>14} {fmt_usd(pen):>14} {fmt_usd(ni + pen):>14}\n")


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #


def select_targets(issuers: list[Issuer], directory: SecDirectory, sec: HttpClient,
                   top_n: int, refresh: bool) -> list[tuple[Issuer, AnnualFiling]]:
    """Walk MSOS issuers by weight and keep the first `top_n` that file an audited
    annual report on EDGAR; record why any issuer was skipped."""
    selected: list[tuple[Issuer, AnnualFiling]] = []
    for iss in issuers:
        if len(selected) >= top_n:
            break
        entity = directory.resolve(iss)
        if entity is None:
            iss.status = "skipped: not an SEC registrant (no CIK match)"
            log.info("Skip %-30s %s", iss.name, iss.status)
            continue
        iss.cik, iss.sec_ticker, iss.sec_name = entity.cik, entity.ticker, entity.title
        try:
            filing = find_latest_annual_filing(sec, entity.cik, refresh)
        except PipelineError as exc:
            iss.status = f"skipped: submissions lookup failed ({exc})"
            continue
        if filing is None:
            iss.status = "skipped: no 10-K/40-F/20-F on EDGAR"
            log.info("Skip %-30s %s", iss.name, iss.status)
            continue
        filing.ticker = filing.ticker or entity.ticker
        iss.status = f"selected: {filing.form} filed {filing.filing_date}"
        log.info("Target %d: %-28s %-6s CIK %-8d %s %s (period %s)", len(selected) + 1, iss.name,
                 filing.ticker, filing.cik, filing.form, filing.accession, filing.report_date)
        selected.append((iss, filing))
    return selected


def issuers_from_tickers(tickers: list[str], directory: SecDirectory) -> list[Issuer]:
    issuers = []
    for t in tickers:
        hit = directory.by_ticker.get(t.upper())
        if hit is None:
            raise FatalError(f"Ticker {t} is not in SEC company_tickers.json")
        issuers.append(Issuer(key=issuer_key(hit.title), name=hit.title, weight_pct=0.0,
                              market_value=0.0, tickers=[t.upper()], components=["--tickers override"]))
    return issuers


def run(args: argparse.Namespace) -> int:
    ua = args.sec_user_agent or os.environ.get("SEC_USER_AGENT", "")
    if "@" not in ua:
        raise FatalError("Set SEC_USER_AGENT (or --sec-user-agent) to 'Firm Name contact@firm.com'; "
                         "the SEC blocks undeclared automated traffic.")
    cache = Path(args.cache_dir)
    out_dir = Path(args.output_dir)
    sec = HttpClient(ua, max_rps=args.sec_max_rps, cache_dir=cache / "sec")
    web = HttpClient(WEB_USER_AGENT, max_rps=1.0, cache_dir=cache / "web")

    directory = SecDirectory.load(sec, args.refresh)
    if args.tickers:
        issuers = issuers_from_tickers([t.strip() for t in args.tickers.split(",") if t.strip()], directory)
        holdings_source, as_of = "user-supplied --tickers", None
    else:
        issuers, holdings_source, as_of = fetch_msos_issuers(web, args.holdings_file, args.refresh)

    targets = select_targets(issuers, directory, sec, args.top_n, args.refresh)
    if not targets:
        raise FatalError("No MSOS issuer resolved to an SEC annual report")

    extractor = None if args.dry_run else LLMExtractor(model=args.model, effort=args.effort, cache_dir=cache / "llm")
    footnote_dir = out_dir / "footnotes"
    footnote_dir.mkdir(parents=True, exist_ok=True)

    results: list[CompanyResult] = []
    for rank, (iss, filing) in enumerate(targets, 1):
        res = CompanyResult(rank=rank, issuer=iss.name, msos_weight_pct=iss.weight_pct,
                            company=filing.company, ticker=filing.ticker, cik=filing.cik, form=filing.form,
                            accession=filing.accession, fiscal_year_end=filing.report_date,
                            filing_url=filing.url)
        results.append(res)
        log.info("[%d/%d] %s %s %s", rank, len(targets), filing.company, filing.form, filing.accession)
        try:
            sections = locate_financial_sections(sec, filing)
            res.document_url, res.footnote_method = sections.document_url, sections.tax_note_method
            stem = f"{rank:02d}_{filing.ticker or filing.cik}"
            (footnote_dir / f"{stem}_income_taxes.txt").write_text(sections.tax_note)
            if sections.income_statement:
                (footnote_dir / f"{stem}_statement_of_operations.txt").write_text(sections.income_statement)
            if extractor is None:
                res.status = "dry-run"
                continue
            extraction = extractor.extract(filing, filing.company, sections)
            (footnote_dir / f"{stem}_extraction.json").write_text(extraction.model_dump_json(indent=2))
            apply_extraction(res, extraction, narrative_reserves=args.narrative_reserves)
            try:
                period = filing.report_date or res.fiscal_year_end
                apply_xbrl_check(res, xbrl_net_income(sec, filing, period))
            except (PipelineError, ValueError) as exc:
                res.xbrl_check = f"XBRL check unavailable ({exc})"
            res.status = "ok"
        except PipelineError as exc:
            res.status, res.error = "error", str(exc)
            log.error("  %s: %s", filing.company, exc)

    if args.dry_run:
        log.info("Dry run complete: footnotes written to %s (no LLM calls made)", footnote_dir)
        return 0

    summary = build_executive_summary(results, holdings_source, as_of)
    write_outputs(out_dir, results, issuers, summary, holdings_source, as_of, args.model)
    print_table(results)
    print(summary + "\n")
    log.info("Report written to %s", out_dir / "report.md")
    return 0 if all(r.status == "ok" for r in results) else 2


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0], formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--top-n", type=int, default=5, help="number of SEC-reporting MSOS operators to analyse (default 5)")
    p.add_argument("--tickers", help="comma-separated OTC tickers to analyse instead of the MSOS holdings file")
    p.add_argument("--holdings-file", help="local copy of the MSOS holdings XLSX/CSV (skips the download)")
    p.add_argument("--model", default=DEFAULT_MODEL, help=f"Claude model (default {DEFAULT_MODEL})")
    p.add_argument("--effort", default=DEFAULT_EFFORT, choices=["low", "medium", "high", "xhigh", "max"])
    p.add_argument("--sec-user-agent", help="overrides SEC_USER_AGENT")
    p.add_argument("--sec-max-rps", type=float, default=5.0, help="SEC request ceiling; SEC allows 10/s (default 5)")
    p.add_argument("--output-dir", default="output")
    p.add_argument("--cache-dir", default=".cache")
    p.add_argument("--refresh", action="store_true", help="ignore cached holdings/submissions data")
    p.add_argument("--dry-run", action="store_true", help="run every stage except the LLM call")
    p.add_argument("--narrative-reserves", action="store_true",
                   help="also count uncertain-tax-position reserve lines the footnote ties to 280E only in "
                        "narrative, without attributing an amount to 280E")
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    if load_dotenv is not None:
        load_dotenv()
    args = parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    for noisy in ("urllib3", "httpx", "httpx2", "anthropic"):
        logging.getLogger(noisy).setLevel(logging.WARNING)
    try:
        return run(args)
    except FatalError as exc:
        log.error("FATAL: %s", exc)
        return 1
    except PipelineError as exc:
        log.error("Pipeline aborted: %s", exc)
        return 1
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())
