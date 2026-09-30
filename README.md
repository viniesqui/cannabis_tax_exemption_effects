# MSOS Section 280E Diligence Pipeline

`msos_280e_diligence.py` quantifies the Section 280E tax penalty in the audited
financial statements of the largest U.S. multi-state cannabis operators (MSOs)
and restates their net income as if cannabis moved to Schedule III.

Financial-data vendors report total income tax expense, but they do not report
the 280E component. That number exists only in each company's income-tax
footnote, inside the statutory-to-effective tax rate reconciliation. This
pipeline goes to the primary source to get it.

## Pipeline

| Stage | Source | What happens |
|---|---|---|
| 1. Target identification | [AdvisorShares MSOS holdings file](https://advisorshares.com/wp-content/uploads/csv/holdings/AdvisorShares_MSOS_Holdings_File.xlsx) | Header-agnostic XLSX/CSV parser; total-return-swap and direct-equity lines are aggregated by issuer; cash and derivative collateral are excluded |
| 2. Ticker / CIK resolution | [SEC `company_tickers.json`](https://www.sec.gov/files/company_tickers.json) | Name/ticker matching resolves OTC tickers (CURLF, GTBIF, TCNNF, VRNOF, …) to CIKs |
| 3. Filing retrieval | [EDGAR submissions API](https://www.sec.gov/search-filings/edgar-application-programming-interfaces) | Most recent audited annual report: **10-K**, or **40-F** for Canadian-domiciled MSOs (Curaleaf, Cresco Labs, Glass House). Issuers with no EDGAR annual report are skipped and the next-largest holding is used |
| 4. Footnote isolation | Inline XBRL filing | Extracts the issuer's own `us-gaap:IncomeTaxDisclosureTextBlock` tag, following `continuedAt` chains. Falls back to a scored heading search when the tag is missing. Also locates the consolidated statement of operations |
| 5. LLM parsing | Anthropic API (`claude-opus-5-5`) | A strict JSON schema (structured outputs) makes Claude transcribe every reconciliation line with its label, amount, percentage and category. Claude never computes anything |
| 6. Financial logic | deterministic Python | Selects the 280E line(s), or the uncertain-tax-position reserve line the footnote attributes to 280E, applies the unit scale, computes **Pro Forma NI = Reported NI + 280E penalty** and effective tax rates |
| 7. Controls and output | [XBRL company facts](https://data.sec.gov/api/xbrl/companyfacts/CIK0001754195.json) | Checks that reconciliation lines sum to total tax expense and that net income ties to the XBRL `ProfitLoss`/`NetIncomeLoss` facts. Writes JSON, CSV and a Markdown report whose executive summary is generated from the computed figures, with a source link for each filing |

FY2025 is the first year public companies must follow
[ASU 2023-09](https://dart.deloitte.com/USDART/home/publications/deloitte/heads-up/2025/income-tax-disclosure-considerations-related-adoption-asu-2023-09).
It requires rate reconciliations in both dollars and percentages, and any
nondeductible item of 5% or more must be shown on its own line. 280E is far
above that threshold for every MSO, so it now appears as its own dollar line.

## Setup

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

### Credentials (keep them out of code, chat and git)

The script reads two environment variables:

| Variable | Purpose |
|---|---|
| `MSOS_ANTHROPIC_API_KEY` | Anthropic API key for the extraction calls. `ANTHROPIC_API_KEY` works as a fallback, but the project-specific name keeps the key from colliding with Claude Code's own login in a cloud session. The script never logs it or writes it to disk. |
| `SEC_USER_AGENT` | `"Firm Name contact@firm.com"`. The [SEC fair-access policy](https://www.sec.gov/os/accessing-edgar-data) requires a declared user agent. Requests without one are blocked. |

Safe ways to provide them:

- **Claude Code on the web:** at [claude.ai/code](https://claude.ai/code), select
  the cloud button showing the environment name, in the row above the message
  box. Choose **Add cloud environment**, or hover over an existing environment
  and select its settings (gear) icon. In the dialog:
  - set **Network access** to **Custom**;
  - list `www.sec.gov`, `data.sec.gov` and `advisorshares.com` under **Allowed
    domains**, and keep **Also include default list of common package
    managers** checked;
  - add both variables under **Environment variables**, one `KEY=value` per line.

  Variables are copied into a session when it starts, so start a new session
  afterwards. Don't use the dialog's **API credentials** section for the
  Anthropic key: the agent proxy never attaches stored credentials to
  `api.anthropic.com`. Environment-variable values are readable by anyone who
  uses the environment, so keep the key in a personal (not organization-shared)
  environment dedicated to this project.
- **Local shell:** `export MSOS_ANTHROPIC_API_KEY=...` in your terminal. Or copy
  `.env.example` to `.env`, which is git-ignored and loaded automatically.

Use a dedicated key with a spend limit, created in the Anthropic Console, and
revoke it when the project is done. A full run makes 5 Claude requests, which
typically costs a few dollars at most. Results are cached, so re-runs cost
nothing.

## Run

```bash
python msos_280e_diligence.py --dry-run   # all stages except the LLM; writes isolated footnotes for review
python msos_280e_diligence.py             # full run, top 5 SEC-reporting MSOS operators
python msos_280e_diligence.py --top-n 7 --effort xhigh
python msos_280e_diligence.py --no-narrative-reserves   # carry reserves tied to 280E only in narrative at zero
python msos_280e_diligence.py --tickers TCNNF,GTBIF,CURLF,VRNOF,CRLBF      # bypass the holdings file
python msos_280e_diligence.py --holdings-file ~/Downloads/MSOS.xlsx        # use a manually downloaded file
```

Outputs are written to `output/`:

- `report.md`: results table, the 280E lines used with quotes, target selection, methodology.
- `executive_summary.md`: a one-paragraph summary with source links.
- `results.csv` and `results.json`.
- `footnotes/`: the exact footnote text and statement of operations sent to the LLM, plus the raw JSON extraction for each company. This is the audit trail.

## Robustness

- **SEC rate limits.** Requests are capped at 5/s; SEC allows 10/s. The script
  honours `Retry-After` and backs off exponentially on 429/5xx and connection
  errors. On SEC's `403 Request Rate Threshold Exceeded` it cools off for 60s+
  and halves its own request rate for the rest of the run. An
  undeclared-user-agent response or a proxy/TLS denial stops the run
  immediately with a fix-it message instead of retrying. EDGAR archive
  documents are cached on disk, since they never change.
- **LLM calls.** The SDK retries 429/5xx/connection errors. An outer retry
  layer adds longer, `retry-after`-aware backoff. Authentication, permission
  and bad-model errors abort the run. Refusals and truncation are detected
  from `stop_reason`. Server-side refusal fallback (`fallbacks: "default"`) is
  enabled. Responses are validated with Pydantic and cached by
  prompt/model/effort hash.
- **Failure isolation.** A failure for one company is recorded in the report
  with its reason, and the rest of the run continues.

## Methodology and limitations

- **Pro forma logic.** Section 280E disallows every deduction except cost of
  goods sold for a business trafficking in a Schedule I or II substance. The
  reconciliation line measures exactly that: the tax effect of the disallowed
  deductions. Rescheduling to Schedule III removes the disallowance, so this
  permanent difference disappears and tax expense falls by the same amount.
  Revenue, gross margin and operating costs are unchanged.
- **Line selection.** Code picks the first of these that applies:
  1. Lines the filing explicitly attributes to 280E.
  2. Uncertain-tax-position reserve lines. Most large MSOs now file their
     returns as if 280E does not apply and reserve for the disputed tax, so in
     FY2025 the 280E cost sits in the change in that reserve, not in a
     nondeductible line. A reserve line counts when the footnote attributes
     that line to 280E (for example a footnote marker on it) or identifies the
     reserve as the company's 280E position (for example a rollforward row
     labelled 280E). Lines that hold only interest and penalties are excluded.
     Reserves the footnote ties to 280E only in general narrative (for
     example Green Thumb's "this results in unrecognized tax benefits") also
     count, unless you pass `--no-narrative-reserves`. Flags:
     `280E_FROM_RESERVE_LINE`, `RESERVE_INCLUDES_INTEREST_PENALTIES`,
     `RESERVE_NARRATIVE_TIE_ONLY`.
  3. Nothing, when the company reserves for 280E but no reserve line
     qualifies. Its penalty is carried at zero and flagged
     `280E_IN_RESERVE_NOT_QUANTIFIED`.
  4. For companies with no 280E reserve, every nondeductible line, used as a
     proxy only when the footnote discusses 280E and flagged
     `280E_NOT_SEPARATELY_LABELLED`. This proxy can include unrelated permanent
     differences, so check the flagged rows.
- **Reserve lines are approximate.** A reserve line records the change in the
  reserve for the year. It can include interest, penalties and positions for
  prior years, so it can differ from the current year's 280E cost.
- **Not modelled:** state conformity to 280E, deferred-tax remeasurement,
  release of reserves accrued in prior years, the timing of cash tax, and
  second-order effects on price competition.
- **Regulatory status (September 2026).** In April 2026 the Department of
  Justice moved FDA-approved marijuana products and state-licensed medical
  marijuana to Schedule III
  ([Federal Register 2026-08177](https://www.federalregister.gov/documents/2026/04/28/2026-08177/schedules-of-controlled-substances-rescheduling-of-marijuana)).
  Adult-use cannabis remains in Schedule I. The DEA's hearing on broader
  rescheduling concluded in July 2026 and a decision is pending
  ([DEA](https://www.dea.gov/marijuana-rescheduling-regulatory-actions)).
  The pro forma therefore represents full rescheduling. The medical-only relief
  already granted captures part of it.

## Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

The suite runs fully offline. Synthetic inline-XBRL 10-K/40-F documents,
EDGAR JSON payloads, an MSOS holdings workbook and a fake Claude client cover
the following:

- footnote isolation, including the fallback heuristic's rejection of MD&A and the table of contents;
- swap/equity aggregation;
- CIK resolution;
- 280E selection edge cases: percent-only tables, sign normalisation and proxy lines;
- tie-out controls;
- HTTP rate-limit and retry behaviour;
- LLM error classification and caching;
- an end-to-end run.
