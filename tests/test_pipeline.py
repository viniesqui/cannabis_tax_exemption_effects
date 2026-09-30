"""Offline tests for msos_280e_diligence.py.

Every network call is replaced by fixtures shaped like the real endpoints
(EDGAR submissions/companyfacts JSON, inline-XBRL 10-K/40-F documents, the
AdvisorShares holdings workbook) and the Claude client is faked, so the suite
runs without SEC access or an API key.
"""

from __future__ import annotations

import io
import json
import sys
from pathlib import Path
from types import SimpleNamespace

import openpyxl
import pytest
import requests

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import msos_280e_diligence as m  # noqa: E402

# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #

TENK_HTML = """<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:ix="http://www.xbrl.org/2013/inlineXBRL">
<head><title>tcnnf-20251231</title></head>
<body>
<div style="display:none"><ix:header><ix:resources><div>HIDDEN-XBRL-CONTEXT</div></ix:resources></ix:header></div>
<p>TABLE OF CONTENTS</p>
<table><tr><td>Income Taxes</td><td>F-30</td></tr></table>
<p>Item 7. Management's Discussion and Analysis</p>
<p>Results of Operations</p>
<table>
<tr><td>Revenue</td><td>$</td><td>1,200,000</td></tr>
<tr><td>Income tax expense</td><td>$</td><td>200,000</td></tr>
<tr><td>Net loss</td><td>$</td><td>(50,000</td><td>)</td></tr>
</table>
<p>Income Taxes</p>
<p>Our effective tax rate is driven by Section 280E; the statutory rate is 21%.</p>
<p>CONSOLIDATED STATEMENTS OF OPERATIONS</p>
<p>(in thousands of U.S. dollars, except per share data)</p>
<table>
<tr><td></td><td>2025</td><td>2024</td></tr>
<tr><td>Revenue</td><td>$</td><td>1,200,000</td><td>$</td><td>1,150,000</td></tr>
<tr><td>Gross profit</td><td></td><td>700,000</td><td></td><td>650,000</td></tr>
<tr><td>Income before provision for income taxes</td><td></td><td>150,000</td><td></td><td>120,000</td></tr>
<tr><td>Provision for income taxes</td><td></td><td>200,000</td><td></td><td>190,000</td></tr>
<tr><td>Net loss</td><td>$</td><td>(50,000</td><td>)</td><td>$</td><td>(70,000</td><td>)</td></tr>
<tr><td>Net loss per share - basic and diluted</td><td>$</td><td>(0.26</td><td>)</td><td>$</td><td>(0.37</td><td>)</td></tr>
</table>
<p>NOTES TO CONSOLIDATED FINANCIAL STATEMENTS</p>
<ix:nonNumeric name="us-gaap:IncomeTaxDisclosureTextBlock" contextRef="c1" continuedAt="cont1" escape="true">
<p>15. INCOME TAXES</p>
<p>As the Company operates in the cannabis industry, it is subject to the limits of IRC Section 280E.</p>
<table>
<tr><td></td><td>2025</td><td></td></tr>
<tr><td>Expected income tax expense at federal statutory rate</td><td>$</td><td>31,500</td><td>21.0</td><td>%</td></tr>
<tr><td>Nondeductible expenses - IRC Section 280E</td><td></td><td>150,000</td><td>100.0</td><td>%</td></tr>
<tr><td>State taxes, net of federal benefit</td><td></td><td>18,500</td><td>12.3</td><td>%</td></tr>
<tr><td>Total income tax expense</td><td>$</td><td>200,000</td><td>133.3</td><td>%</td></tr>
</table>
</ix:nonNumeric>
<div>F-31</div>
<ix:continuation id="cont1"><p>Deferred tax assets are reduced by a valuation allowance. Unrecognized tax benefits relate to 280E.</p></ix:continuation>
<p>16. COMMITMENTS AND CONTINGENCIES</p>
<p>The Company is party to various legal proceedings.</p>
</body></html>
"""


def plain_html(html: str) -> str:
    """The same filing without inline-XBRL tags (forces the heading heuristic)."""
    for tag in ("ix:nonNumeric", "ix:continuation"):
        html = html.replace(f"<{tag}", "<div").replace(f"</{tag}>", "</div>")
    return html


def make_xlsx(rows: list[list]) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


HOLDINGS_ROWS = [
    ["AdvisorShares Pure US Cannabis ETF"],
    ["Holdings as of 09/29/2026"],
    [],
    ["Fund Name", "Security Name", "Ticker", "CUSIP", "Shares", "Market Value", "% of Net Assets"],
    ["MSOS", "Derivatives Collateral Nomura", "", "", 0, 330_000_000, 0.3368],
    ["MSOS", "Curaleaf Holdings Inc", "CURLF", "X", 1, 117_000_000, 0.1194],
    ["MSOS", "Trulieve Cannabis Corp", "TCNNF", "X", 1, 98_000_000, 0.0999],
    ["MSOS", "Curaleaf Holdings Inc Swap", "", "", 0, 95_000_000, 0.0972],
    ["MSOS", "Green Thumb Industries Swap", "", "", 0, 83_000_000, 0.085],
    ["MSOS", "Green Thumb Industries Inc", "GTBIF", "X", 1, 20_000_000, 0.02],
    ["MSOS", "Verano Holdings Corp Swap", "", "", 0, 49_000_000, 0.05],
    ["MSOS", "Cresco Labs Inc Swap", "", "", 0, 44_000_000, 0.045],
    ["MSOS", "Glass House Brands Swap", "", "", 0, 39_000_000, 0.04],
    ["MSOS", "Dreyfus Government Cash Management", "", "", 0, 29_000_000, 0.03],
    ["MSOS", "Cash & Other", "", "", 0, 10_000_000, 0.01],
]

SEC_TICKERS = {
    "0": {"cik_str": 1756770, "ticker": "CURLF", "title": "Curaleaf Holdings, Inc."},
    "1": {"cik_str": 1754195, "ticker": "TCNNF", "title": "Trulieve Cannabis Corp."},
    "2": {"cik_str": 1795139, "ticker": "GTBIF", "title": "Green Thumb Industries Inc."},
    "3": {"cik_str": 1848416, "ticker": "VRNOF", "title": "Verano Holdings Corp."},
    "4": {"cik_str": 9999999, "ticker": "CURA", "title": "Cura Partners Inc"},
}


def submissions(name, ticker, rows):
    cols = {"accessionNumber": [], "filingDate": [], "reportDate": [], "form": [], "primaryDocument": []}
    for acc, fdate, rdate, form, doc in rows:
        for k, v in zip(cols, (acc, fdate, rdate, form, doc), strict=True):
            cols[k].append(v)
    return {"name": name, "tickers": [ticker], "filings": {"recent": cols}}


def extraction_json(net_income=-50_000, penalty=150_000, category="section_280e", quotes=True,
                    amount=True, percent=100.0, pretax=150_000, total=200_000, statutory=31_500, state=18_500):
    line = {"label": "Nondeductible expenses - IRC Section 280E", "category": category,
            "amount": penalty if amount else None, "percent": percent,
            "increases_tax_expense": True, "rationale": "Label cites IRC Section 280E."}
    return {
        "fiscal_year_end": "2025-12-31",
        "income_statement": {"unit_multiplier": 1000, "net_income_label": "Net loss", "net_income": net_income,
                             "pretax_income": pretax, "income_tax_expense": total, "total_revenue": 1_200_000},
        "rate_reconciliation": {
            "found": True, "unit_multiplier": 1000,
            "presentation": "amounts_and_percentages" if amount else "percentages_only",
            "statutory_rate_percent": 21.0, "pretax_income": pretax, "total_income_tax_expense": total,
            "line_items": [
                {"label": "Expected income tax expense at federal statutory rate", "category": "statutory",
                 "amount": statutory if amount else None, "percent": 21.0, "increases_tax_expense": True,
                 "rationale": "Starting line."},
                line,
                {"label": "State taxes, net of federal benefit", "category": "state_local",
                 "amount": state if amount else None, "percent": 12.3, "increases_tax_expense": True,
                 "rationale": "State."},
            ],
        },
        "section_280e_quotes": ["it is subject to the limits of IRC Section 280E."] if quotes else [],
        "uncertain_tax_position_on_280e": False,
        "extraction_notes": "",
    }


# --------------------------------------------------------------------------- #
# HTML / footnote isolation
# --------------------------------------------------------------------------- #


def test_render_table_glues_currency_fragments():
    soup = m.parse_html(TENK_HTML)
    stmt = [t for t in soup.find_all("table") if "Gross profit" in t.get_text()][0]
    rendered = m.render_table(stmt)
    assert "Net loss | $(50,000) | $(70,000)" in rendered
    assert "Revenue | $1,200,000 | $1,150,000" in rendered


def test_hidden_ixbrl_header_is_excluded():
    text = m.html_to_text(m.parse_html(TENK_HTML))
    assert "HIDDEN-XBRL-CONTEXT" not in text


def test_ixbrl_text_block_follows_continuations():
    note = m.extract_ixbrl_text_block(m.parse_html(TENK_HTML), "IncomeTaxDisclosureTextBlock")
    assert note.startswith("15. INCOME TAXES")
    assert "Nondeductible expenses - IRC Section 280E | 150,000 | 100.0%" in note
    assert "valuation allowance" in note  # from the ix:continuation
    assert "F-31" not in note and "COMMITMENTS" not in note


def test_heading_heuristic_finds_footnote_not_mdna_or_toc():
    soup = m.parse_html(plain_html(TENK_HTML))
    assert m.extract_ixbrl_text_block(soup, "IncomeTaxDisclosureTextBlock") is None
    note = m.extract_tax_note_heuristic(m.html_to_text(soup))
    assert note is not None and note.startswith("15. INCOME TAXES")
    assert "Section 280E | 150,000" in note
    assert "COMMITMENTS AND CONTINGENCIES" not in note
    assert "Our effective tax rate is driven" not in note  # MD&A discussion rejected


def test_income_statement_prefers_audited_statement_over_mdna_table():
    stmt = m.extract_income_statement(m.parse_html(TENK_HTML))
    assert "CONSOLIDATED STATEMENTS OF OPERATIONS" in stmt
    assert "(in thousands of U.S. dollars" in stmt
    assert "Gross profit" in stmt


def test_rank_filing_documents_for_40f():
    names = ["form40-f.htm", "q425-exx991xaif.htm", "q425-exx999xconsent.htm", "curlf-20251231.htm",
             "q42025-financialstatements.htm", "0001756770-26-000019-index.htm", "R1.htm.xml"]
    ranked = m.rank_filing_documents(names, "form40-f.htm")
    assert ranked[:3] == ["form40-f.htm", "curlf-20251231.htm", "q42025-financialstatements.htm"]
    assert ranked[-2:] == ["q425-exx991xaif.htm", "q425-exx999xconsent.htm"] or set(ranked[-2:]) == {
        "q425-exx991xaif.htm", "q425-exx999xconsent.htm"}
    assert all("index" not in n for n in ranked)


# --------------------------------------------------------------------------- #
# Holdings and resolution
# --------------------------------------------------------------------------- #


def test_holdings_xlsx_aggregates_swaps_and_excludes_collateral():
    holdings, as_of = m.parse_holdings(make_xlsx(HOLDINGS_ROWS))
    assert as_of == "09/29/2026"
    issuers = m.aggregate_by_issuer(holdings)
    names = [i.name for i in issuers]
    assert names[:6] == ["Curaleaf Holdings Inc", "Green Thumb Industries", "Trulieve Cannabis Corp",
                         "Verano Holdings Corp", "Cresco Labs Inc", "Glass House Brands"]
    assert issuers[0].weight_pct == pytest.approx(21.66)
    assert not any(k in " ".join(names).lower() for k in ("collateral", "dreyfus", "cash"))


def test_holdings_csv_with_percent_strings():
    csv_bytes = (b"Date,StockTicker,SecurityName,MarketValue,Weightings\n"
                 b"09/29/2026,TCNNF,Trulieve Cannabis Corp,98000000,9.99%\n"
                 b"09/29/2026,,Trulieve Cannabis Corp Swap,50000000,5.01%\n")
    holdings, as_of = m.parse_holdings(csv_bytes, "h.csv")
    assert as_of == "09/29/2026"
    [iss] = m.aggregate_by_issuer(holdings)
    assert iss.weight_pct == pytest.approx(15.0)
    assert iss.tickers == ["TCNNF"]


def test_sec_resolution_by_name_ticker_guard_and_seed():
    directory = m.SecDirectory([m.SecEntity(int(v["cik_str"]), v["ticker"], v["title"]) for v in SEC_TICKERS.values()])
    cura = m.Issuer(key=m.issuer_key("Curaleaf Holdings Inc Swap"), name="Curaleaf", weight_pct=1,
                    market_value=1, tickers=["CURA"], components=[])
    assert directory.resolve(cura).ticker == "CURLF"  # 'CURA' belongs to another company
    gti = m.Issuer(key=m.issuer_key("Green Thumb Industries Swap"), name="GTI", weight_pct=1,
                   market_value=1, tickers=[], components=[])
    assert directory.resolve(gti).cik == 1795139
    cresco = m.Issuer(key=m.issuer_key("Cresco Labs Inc Swap"), name="Cresco", weight_pct=1,
                      market_value=1, tickers=[], components=[])
    assert directory.resolve(cresco).cik == 1832928  # seed fallback
    unknown = m.Issuer(key="acme weed", name="Acme", weight_pct=1, market_value=1, tickers=[], components=[])
    assert directory.resolve(unknown) is None


class FakeHttp:
    def __init__(self, routes: dict[str, bytes | dict | list]):
        self.routes = routes
        self.requested: list[str] = []

    def get(self, url, *, ttl=None, use_cache=True):
        self.requested.append(url)
        if url not in self.routes:
            raise m.HttpNotFound(f"404 Not Found: {url}")
        body = self.routes[url]
        return body if isinstance(body, bytes) else json.dumps(body).encode()

    def get_json(self, url, *, ttl=None, use_cache=True):
        return json.loads(self.get(url, ttl=ttl, use_cache=use_cache))


def test_latest_annual_filing_skips_amendments_and_interims():
    sub = submissions("Curaleaf Holdings, Inc.", "CURLF", [
        ("0001756770-26-000090", "2026-09-02", "2026-09-02", "8-K", "a.htm"),
        ("0001756770-26-000050", "2026-05-01", "2025-12-31", "10-K/A", "b.htm"),
        ("0001756770-26-000019", "2026-03-10", "2025-12-31", "40-F", "form40-f.htm"),
        ("0001756770-25-000016", "2025-03-05", "2024-12-31", "40-F", "old.htm"),
    ])
    http = FakeHttp({m.SEC_SUBMISSIONS_URL.format(cik=1756770): sub})
    f = m.find_latest_annual_filing(http, 1756770, refresh=False)
    assert (f.form, f.accession, f.report_date) == ("40-F", "0001756770-26-000019", "2025-12-31")
    assert f.url == "https://www.sec.gov/Archives/edgar/data/1756770/000175677026000019/form40-f.htm"


# --------------------------------------------------------------------------- #
# Financial logic
# --------------------------------------------------------------------------- #


def test_explicit_280e_line_scaled_to_dollars():
    x = m.TaxExtraction.model_validate(extraction_json())
    sel = m.select_280e_penalty(x)
    assert sel.method == "explicit_280e"
    assert sel.amount_usd == 150_000_000
    assert m.reconciliation_ties(x) is True


def test_nondeductible_proxy_only_when_280e_discussed():
    x = m.TaxExtraction.model_validate(extraction_json(category="nondeductible_other"))
    sel = m.select_280e_penalty(x)
    assert sel.method == "nondeductible_proxy" and "280E_NOT_SEPARATELY_LABELLED" in sel.flags
    x2 = m.TaxExtraction.model_validate(extraction_json(category="nondeductible_other", quotes=False))
    sel2 = m.select_280e_penalty(x2)
    assert sel2.amount_usd is None and sel2.method == "not_disclosed"


def test_percent_only_reconciliation_for_loss_company():
    x = m.TaxExtraction.model_validate(extraction_json(amount=False, percent=-120.0, pretax=-100_000))
    sel = m.select_280e_penalty(x)
    assert sel.amount_usd == pytest.approx(120_000_000)
    assert "PERCENT_DERIVED" in sel.flags
    assert m.reconciliation_ties(x) is None


def test_negative_280e_is_sign_normalised_and_flagged():
    x = m.TaxExtraction.model_validate(extraction_json(penalty=-150_000, total=-100_000))
    sel = m.select_280e_penalty(x)
    assert sel.amount_usd == 150_000_000 and "SIGN_NORMALISED" in sel.flags


def test_reconciliation_that_does_not_tie_is_flagged():
    x = m.TaxExtraction.model_validate(extraction_json(state=90_000))
    res = m.CompanyResult(rank=1, issuer="T", msos_weight_pct=10)
    m.apply_extraction(res, x)
    assert "RECONCILIATION_DOES_NOT_TIE" in res.flags


def test_pro_forma_net_income_and_tax_rates():
    res = m.CompanyResult(rank=1, issuer="Trulieve", msos_weight_pct=10)
    m.apply_extraction(res, m.TaxExtraction.model_validate(extraction_json()))
    assert res.reported_net_income == -50_000_000
    assert res.penalty_280e == 150_000_000
    assert res.pro_forma_net_income == 100_000_000
    assert res.effective_tax_rate == pytest.approx(200 / 150)
    assert res.pro_forma_effective_tax_rate == pytest.approx(50 / 150)
    assert res.flags == []


def test_xbrl_cross_check():
    filing = m.AnnualFiling(1754195, "Trulieve", "TCNNF", "10-K", "0001754195-26-000019",
                            "2026-02-26", "2025-12-31", "tcnnf-20251231.htm")
    facts = {"facts": {"us-gaap": {"ProfitLoss": {"units": {"USD": [
        {"start": "2024-01-01", "end": "2024-12-31", "val": -70_000_000, "accn": "0001754195-26-000019"},
        {"start": "2025-10-01", "end": "2025-12-31", "val": -5_000_000, "accn": "0001754195-26-000019"},
        {"start": "2025-01-01", "end": "2025-12-31", "val": -50_000_000, "accn": "0001754195-26-000019"},
    ]}}}}}
    http = FakeHttp({m.SEC_COMPANYFACTS_URL.format(cik=1754195): facts})
    got = m.xbrl_net_income(http, filing, "2025-12-31")
    assert got == {"ProfitLoss": -50_000_000}
    res = m.CompanyResult(rank=1, issuer="T", msos_weight_pct=1, reported_net_income=-50_000_000)
    m.apply_xbrl_check(res, got)
    assert res.xbrl_check == "ties to us-gaap:ProfitLoss"
    res2 = m.CompanyResult(rank=1, issuer="T", msos_weight_pct=1, reported_net_income=-40_000_000)
    m.apply_xbrl_check(res2, got)
    assert "NET_INCOME_XBRL_MISMATCH" in res2.flags


def _ok_result(name, ticker, ni, pen, rev=1e9, weight=10.0):
    return m.CompanyResult(rank=1, issuer=name, msos_weight_pct=weight, company=name, ticker=ticker,
                           form="10-K", fiscal_year_end="2025-12-31", reported_net_income=ni,
                           penalty_280e=pen, pro_forma_net_income=ni + pen, total_revenue=rev,
                           filing_url=f"https://www.sec.gov/{ticker}.htm", status="ok", penalty_method="explicit_280e")


def test_executive_summary_loss_to_profit():
    results = [_ok_result("Alpha", "AAA", -300e6, 350e6), _ok_result("Beta", "BBB", 50e6, 150e6)]
    text = m.build_executive_summary(results, m.MSOS_HOLDINGS_URL, "09/29/2026")
    assert "aggregate Section 280E penalty of $500.0M" in text
    assert "from an aggregate net loss of $250.0M to a pro forma net profit of $250.0M" in text
    assert "profitable operators rising from 1 to 2 of 2" in text
    assert "(Alpha (AAA) and Beta (BBB); together 20.0% of MSOS net assets" in text
    assert "Across the two largest" in text
    assert f"({m.MSOS_HOLDINGS_URL})" in text and "(https://www.sec.gov/AAA.htm)" in text


def test_executive_summary_narrowing_loss_and_ticker_override():
    results = [_ok_result("Alpha", "AAA", -300e6, 100e6, weight=0.0)]
    text = m.build_executive_summary(results, "user-supplied --tickers", None)
    assert "aggregate net loss narrows from $300.0M to $200.0M, a 33% reduction" in text
    assert "0 of 1 operators profitable pro forma, unchanged from reported" in text
    assert "MSOS net assets" not in text and "user-supplied" not in text


# --------------------------------------------------------------------------- #
# HTTP resilience
# --------------------------------------------------------------------------- #


class FakeResponse:
    def __init__(self, status, body=b"", headers=None):
        self.status_code, self.content, self.headers = status, body, headers or {}
        self.text = body.decode()


class FakeSession:
    def __init__(self, script):
        self.script, self.headers, self.calls = list(script), {}, 0

    def get(self, url, timeout):
        self.calls += 1
        item = self.script.pop(0)
        if isinstance(item, Exception):
            raise item
        return item


def client(script, **kw):
    sleeps: list[float] = []
    c = m.HttpClient("Test Co test@example.com", max_rps=1000, cache_dir=None, session=FakeSession(script),
                     sleep=sleeps.append, **kw)
    return c, sleeps


def test_retry_after_on_429():
    c, sleeps = client([FakeResponse(429, b"slow down", {"Retry-After": "7"}), FakeResponse(200, b"ok")],
                       rate_limit_cooldown=0)
    assert c.get("https://data.sec.gov/x") == b"ok"
    assert 7 in sleeps


def test_sec_rate_threshold_403_backs_off_and_slows_down():
    body = b"<html>Request Rate Threshold Exceeded</html>"
    c, sleeps = client([FakeResponse(403, body), FakeResponse(200, b"ok")])
    before = c.min_interval
    assert c.get("https://www.sec.gov/x") == b"ok"
    assert max(sleeps) >= 60 and c.min_interval == pytest.approx(before * 2)


def test_undeclared_user_agent_is_fatal():
    c, _ = client([FakeResponse(403, b"Your Request Originates from an Undeclared Automated Tool")])
    with pytest.raises(m.FatalError):
        c.get("https://www.sec.gov/x")


def test_connection_errors_retry_then_exhaust():
    c, sleeps = client([requests.ConnectionError("reset"), FakeResponse(503), FakeResponse(200, b"ok")])
    assert c.get("https://www.sec.gov/x") == b"ok" and len(sleeps) >= 2
    c2, _ = client([FakeResponse(500)] * 3, max_retries=2)
    with pytest.raises(m.PipelineError, match="after 3 attempts"):
        c2.get("https://www.sec.gov/x")
    c3, _ = client([FakeResponse(404)])
    with pytest.raises(m.HttpNotFound):
        c3.get("https://www.sec.gov/x")


def test_proxy_denial_fails_fast():
    c, sleeps = client([requests.exceptions.ProxyError("CONNECT tunnel failed, response 403")])
    with pytest.raises(m.FatalError, match="www.sec.gov"):
        c.get("https://www.sec.gov/x")
    assert sleeps == []


def test_disk_cache(tmp_path):
    session = FakeSession([FakeResponse(200, b"payload")])
    c = m.HttpClient("T t@x.com", max_rps=1000, cache_dir=tmp_path, session=session, sleep=lambda s: None)
    assert c.get("https://www.sec.gov/a") == b"payload"
    assert c.get("https://www.sec.gov/a") == b"payload"
    assert session.calls == 1


# --------------------------------------------------------------------------- #
# LLM layer
# --------------------------------------------------------------------------- #


class FakeStream:
    def __init__(self, message):
        self.message = message

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def get_final_message(self):
        if isinstance(self.message, Exception):
            raise self.message
        return self.message


class FakeMessages:
    def __init__(self, responder):
        self.responder, self.calls = responder, []

    def stream(self, **kwargs):
        self.calls.append(kwargs)
        return FakeStream(self.responder(kwargs))


def fake_client(responder):
    return SimpleNamespace(beta=SimpleNamespace(messages=FakeMessages(responder)))


def message(payload, stop_reason="end_turn", extra_blocks=()):
    blocks = [*extra_blocks, SimpleNamespace(type="thinking", thinking=""),
              SimpleNamespace(type="text", text=json.dumps(payload))]
    return SimpleNamespace(stop_reason=stop_reason, content=blocks, usage=None, stop_details=None)


SAMPLE_FILING = m.AnnualFiling(1754195, "Trulieve Cannabis Corp.", "TCNNF", "10-K", "0001754195-26-000019",
                               "2026-02-26", "2025-12-31", "tcnnf-20251231.htm")
SAMPLE_SECTIONS = m.FinancialSections("https://www.sec.gov/doc.htm", "15. INCOME TAXES ...", "ixbrl", "Net loss | (50,000)")


def test_llm_request_shape_parse_and_cache(tmp_path):
    fc = fake_client(lambda kw: message(extraction_json()))
    ex = m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=tmp_path, client=fc, sleep=lambda s: None)
    x = ex.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS)
    assert x.rate_reconciliation.line_items[1].category == "section_280e"
    call = fc.beta.messages.calls[0]
    assert call["model"] == "claude-opus-5-5"
    assert call["output_config"]["format"]["type"] == "json_schema"
    assert call["output_config"]["effort"] == "high"
    assert call["fallbacks"] == "default" and call["betas"] == ["server-side-fallback-2026-07-01"]
    prompt = call["messages"][0]["content"]
    assert prompt.index("<income_taxes_footnote>") < prompt.index("Work only with the fiscal year ended 2025-12-31")
    ex.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS)
    assert len(fc.beta.messages.calls) == 1  # second call served from cache


def test_llm_text_before_fallback_block_is_discarded():
    stale = SimpleNamespace(type="text", text='{"partial": ')
    fb = SimpleNamespace(type="fallback")
    fc = fake_client(lambda kw: message(extraction_json(), extra_blocks=(stale, fb)))
    ex = m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None, client=fc, sleep=lambda s: None)
    assert ex.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS).fiscal_year_end == "2025-12-31"


def test_llm_refusal_retries_then_fails():
    fc = fake_client(lambda kw: message({}, stop_reason="refusal"))
    sleeps = []
    ex = m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None, client=fc, sleep=sleeps.append,
                        max_attempts=2)
    with pytest.raises(m.PipelineError, match="declined"):
        ex.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS)
    assert len(fc.beta.messages.calls) == 2 and len(sleeps) == 1


def _api_error(cls, status):
    import httpx2

    resp = httpx2.Response(status, request=httpx2.Request("POST", "https://api.anthropic.com/v1/messages"),
                           headers={"retry-after": "3"})
    return cls("boom", response=resp, body=None)


def test_llm_auth_error_is_fatal_and_rate_limit_is_retried():
    anthropic = pytest.importorskip("anthropic")
    fc = fake_client(lambda kw: _api_error(anthropic.AuthenticationError, 401))
    ex = m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None, client=fc, sleep=lambda s: None)
    with pytest.raises(m.FatalError, match="credentials"):
        ex.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS)

    seq = [_api_error(anthropic.RateLimitError, 429), message(extraction_json())]
    sleeps = []
    fc2 = fake_client(lambda kw: seq.pop(0))
    ex2 = m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None, client=fc2, sleep=sleeps.append)
    assert ex2.extract(SAMPLE_FILING, "Trulieve", SAMPLE_SECTIONS).fiscal_year_end == "2025-12-31"
    assert sleeps == [30.0]


def test_missing_api_key_is_fatal(monkeypatch):
    for var in ("MSOS_ANTHROPIC_API_KEY", "ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN"):
        monkeypatch.delenv(var, raising=False)
    with pytest.raises(m.FatalError, match="MSOS_ANTHROPIC_API_KEY"):
        m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None)


def test_project_specific_api_key_is_passed_explicitly(monkeypatch):
    captured = {}
    monkeypatch.setattr(m.anthropic, "Anthropic", lambda **kw: captured.update(kw) or object())
    monkeypatch.setenv("MSOS_ANTHROPIC_API_KEY", "test-key-project")
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key-generic")
    m.LLMExtractor(model="claude-opus-5-5", effort="high", cache_dir=None)
    assert captured["api_key"] == "test-key-project"


def test_schema_matches_pydantic_models():
    def check(schema, model):
        assert schema["additionalProperties"] is False
        assert set(schema["required"]) == set(schema["properties"]) == set(model.model_fields)
        for name, info in model.model_fields.items():
            ann = info.annotation
            sub = schema["properties"][name]
            if isinstance(ann, type) and issubclass(ann, m.BaseModel):
                check(sub, ann)
            elif sub.get("type") == "array" and sub["items"].get("type") == "object":
                check(sub["items"], ann.__args__[0])

    check(m.EXTRACTION_SCHEMA, m.TaxExtraction)


# --------------------------------------------------------------------------- #
# End to end (all network faked)
# --------------------------------------------------------------------------- #


def test_end_to_end(tmp_path, monkeypatch):
    holdings = tmp_path / "msos.xlsx"
    holdings.write_bytes(make_xlsx(HOLDINGS_ROWS[:4] + [
        ["MSOS", "Acme Weed Swap", "", "", 0, 1, 0.25],           # not an SEC registrant -> skipped
        ["MSOS", "Curaleaf Holdings Inc Swap", "", "", 0, 1, 0.20],  # 40-F filer, FS in an exhibit
        ["MSOS", "Trulieve Cannabis Corp", "TCNNF", "", 1, 1, 0.15],  # 10-K filer
        ["MSOS", "Verano Holdings Corp Swap", "", "", 0, 1, 0.05],   # beyond --top-n 2
    ]))
    cura_folder = "https://www.sec.gov/Archives/edgar/data/1756770/000175677026000019"
    tru_folder = "https://www.sec.gov/Archives/edgar/data/1754195/000175419526000019"
    cura_fs = TENK_HTML.replace("tcnnf-20251231", "curlf-20251231")
    routes = {
        m.SEC_TICKERS_URL: SEC_TICKERS,
        m.SEC_SUBMISSIONS_URL.format(cik=1756770): submissions("Curaleaf Holdings, Inc.", "CURLF", [
            ("0001756770-26-000019", "2026-03-10", "2025-12-31", "40-F", "form40-f.htm")]),
        m.SEC_SUBMISSIONS_URL.format(cik=1754195): submissions("Trulieve Cannabis Corp.", "TCNNF", [
            ("0001754195-26-000019", "2026-02-26", "2025-12-31", "10-K", "tcnnf-20251231.htm")]),
        f"{cura_folder}/index.json": {"directory": {"item": [
            {"name": "form40-f.htm"}, {"name": "q425-exx991xaif.htm"}, {"name": "curlf-20251231.htm"}]}},
        f"{cura_folder}/form40-f.htm": b"<html><body><p>FORM 40-F cover. See exhibits.</p></body></html>",
        f"{cura_folder}/curlf-20251231.htm": cura_fs.encode(),
        f"{tru_folder}/index.json": {"directory": {"item": [{"name": "tcnnf-20251231.htm"}]}},
        f"{tru_folder}/tcnnf-20251231.htm": TENK_HTML.encode(),
        m.SEC_COMPANYFACTS_URL.format(cik=1756770): {"facts": {"us-gaap": {"ProfitLoss": {"units": {"USD": [
            {"start": "2025-01-01", "end": "2025-12-31", "val": -300_000_000, "accn": "0001756770-26-000019"}]}}}}},
        m.SEC_COMPANYFACTS_URL.format(cik=1754195): {"facts": {"us-gaap": {"NetIncomeLoss": {"units": {"USD": [
            {"start": "2025-01-01", "end": "2025-12-31", "val": -50_000_000, "accn": "0001754195-26-000019"}]}}}}},
    }
    http = FakeHttp(routes)
    monkeypatch.setattr(m, "HttpClient", lambda *a, **k: http)

    def responder(kw):
        prompt = kw["messages"][0]["content"]
        if "Curaleaf" in prompt:
            return message(extraction_json(net_income=-300_000, penalty=250_000, total=300_000, statutory=31_500,
                                           state=18_500))
        return message(extraction_json())

    real = m.LLMExtractor
    monkeypatch.setattr(m, "LLMExtractor", lambda **kw: real(client=fake_client(responder), **kw))
    monkeypatch.setenv("SEC_USER_AGENT", "Test Research test@example.com")

    out = tmp_path / "out"
    args = m.parse_args(["--holdings-file", str(holdings), "--top-n", "2", "--output-dir", str(out),
                         "--cache-dir", str(tmp_path / "cache")])
    assert m.run(args) == 0

    results = json.loads((out / "results.json").read_text())
    assert [r["ticker"] for r in results] == ["CURLF", "TCNNF"]
    assert [r["form"] for r in results] == ["40-F", "10-K"]
    cura, tru = results
    assert cura["document_url"].endswith("curlf-20251231.htm")
    assert cura["footnote_method"].startswith("ixbrl")
    assert (cura["reported_net_income"], cura["penalty_280e"], cura["pro_forma_net_income"]) == (
        -300e6, 250e6, -50e6)
    assert (tru["reported_net_income"], tru["pro_forma_net_income"]) == (-50e6, 100e6)
    assert cura["xbrl_check"] == "ties to us-gaap:ProfitLoss" and tru["xbrl_check"] == "ties to us-gaap:NetIncomeLoss"

    summary = (out / "executive_summary.md").read_text()
    assert f"({cura_folder}/curlf-20251231.htm)" in summary  # cites the exhibit holding the footnote
    assert f"({m.MSOS_HOLDINGS_URL})" in summary and "file://" not in summary
    assert "aggregate Section 280E penalty of $400.0M" in summary
    assert "from an aggregate net loss of $350.0M to a pro forma net profit of $50.0M" in summary
    report = (out / "report.md").read_text()
    assert "skipped: not an SEC registrant" in report
    assert (out / "footnotes" / "01_CURLF_income_taxes.txt").read_text().startswith("15. INCOME TAXES")
