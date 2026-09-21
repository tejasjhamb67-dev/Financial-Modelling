# Financial Model Construction Agent

A **reusable, general-purpose system** that transforms company disclosure
documents (annual reports, investor/results presentations, regulatory filings)
into a **professional, historically populated, formula-driven, source-traceable,
internally linked Excel financial model** with a 3-year forecast framework.

It is a **model construction engine, not an investment-decision engine.** The
agent automates the mechanical, historical and diligence-heavy work; the analyst
retains full control of forecast assumptions, valuation assumptions, the
investment thesis and the conclusion.

> **Non-negotiable rule: the agent never invents data.** A value that cannot be
> reliably established from the source material is shown as `N/A`,
> `Not Disclosed`, `Not Found` or `Requires Review` — never guessed. Values that
> can legitimately be calculated are marked `DERIVED` and keep their formula.

---

## What it produces

A genuine `.xlsx` workbook (not a PDF-to-Excel dump) containing:

- **Cover** — model info, scope/responsibility split, colour legend, sheet index,
  and a highlighted *"analyst review required"* list.
- **Sources** — full provenance for every extracted datapoint and a document
  inventory (what / when / where / how / why for each number).
- **Reported Financials** — the raw historical statements exactly as disclosed in
  the source documents, in one sheet, right before Assumptions. Disclosed line
  items are hardcoded inputs; the statement totals are explicit additive formulas
  of their components (`=B10+B11+…`, never `SUM()`, never a hardcoded total), so
  the dump is self-checking. This is the single hardcoded source of historical
  truth: the Income Statement, Balance Sheet and Cash Flow **reference this sheet**
  (green links, e.g. COGS = network + access + licence rows) rather than
  re-hardcoding anything. Supply a `reported_statements` block in the package for a
  line-perfect dump; otherwise the sheet falls back to the source database grouped
  by statement. (Cell comments are not used anywhere — provenance lives on Sources.)
- **Assumptions** — the analyst's control panel; every forecast is driven by a
  yellow input cell here (seeded with neutral hold-last-actual placeholders).
- **Revenue Build** — segment build with explicit additive totals and mix %.
- **Operating Drivers** — company-specific KPIs (capacity, volume, utilisation…).
- **Income Statement / Balance Sheet / Cash Flow** — historicals built from
  reported inputs (blue) and explicit additive construction (black); forecasts
  fully formula-driven off Assumptions and schedules.
- **3-Statement Model** — the integrated layer; every line links into the
  statements/schedules (no duplicated hardcodes).
- **Working Capital / Capex & D&A / Debt** — driver schedules that feed back into
  the statements so the model is genuinely integrated and **balances by
  construction** (cash is the plug from the cash-flow statement).
- **Ratios** — growth, margins, returns, leverage, working capital, cash flow —
  all formula-driven off the statements.
- **Valuation** — DCF, EV/EBITDA cross-check, trading-comps template and (with
  segments) a SOTP scaffold. **Frameworks only** — the agent never picks a
  method, multiple, target price or conclusion.
- **Sensitivities** — WACC × terminal-growth grid that recomputes the DCF value.
- **Checks** — balance-sheet identity, cash-flow tie, revenue and debt
  reconciliations, with live `PASS` / `WARNING` / `ERROR` conditional formatting.
  **Checks are never forced to pass.**

Plus two markdown reports: `*_MODEL_BUILD_REPORT.md` and `*_MODEL_DILIGENCE.md`.

### Colour convention (sell-side standard)

| Colour | Meaning |
|---|---|
| **Blue** font | Reported / hardcoded historical input |
| **Black** font | Formula within the same worksheet |
| **Green** font | Link to another worksheet |
| **Yellow** fill | User assumption (analyst input) |
| **Red** | Errors / failed checks |

Constructive totals use explicit additive formulas (`=B10+B11+B12`), never
`SUM()`, so the construction stays visible.

---

## Two ways to run it

- **As a Python engine (this repo):** the deterministic pipeline described below —
  drive it directly with `main.py`, or via the `/modelme` slash command in Claude Code
  (`.claude/commands/modelme.md`).
- **As a Claude Project (no code):** paste
  [`MODELME_PROJECT_INSTRUCTIONS.md`](MODELME_PROJECT_INSTRUCTIONS.md) into a Claude
  Project's custom instructions, upload the company's annual reports as Project
  knowledge, and type `/modelme <Company>`. Claude then runs the same methodology
  (staged flow, Gates 1–3, never-invent-data rule, sheet architecture, colour/formula
  discipline, review checklist — see the SPECIFICATIONS section in that file) directly
  from the uploaded PDFs. Both paths honour the same non-negotiable: never invent data.

## Installation

```bash
pip install -r requirements.txt
```

`openpyxl` and `PyYAML` are required. `pdfplumber` is only needed for the PDF
ingestion path; `pytest` (and optionally `formulas`) only to run the test suite.

## Usage

### 1. Structured company package (reliable path)

The most reliable input is a *company package* — a YAML/JSON file mirroring the
disclosures with provenance. See `inputs/example_company.yaml`.

```bash
python main.py --package inputs/example_company.yaml
```

### 2. PDF documents

Drop annual-report PDFs into a folder and point the agent at it:

```bash
python main.py --documents inputs/documents/ --company "Acme Ltd"
```

The PDF path extracts candidate datapoints **conservatively** — every candidate
is flagged `REQUIRES_REVIEW` with its page/section preserved, because successful
PDF parsing is not successful financial-data extraction. Review the Sources sheet
and diligence report before relying on the numbers.

Outputs land in `outputs/models/`.

### Analyst workflow

1. Run the agent. 2. Open the workbook. 3. Review historicals, the Checks sheet
and the diligence report. 4. Enter your assumptions in the yellow cells. 5. The
model calculates forecasts, ratios and valuation.

---

## Pipeline (mandatory staged flow with validation gates)

```
documents → ingestion → document understanding → disclosure mapping →
extraction → normalization → SOURCE DATABASE → [GATE 1: data validation] →
model specification → [GATE 2: model spec] → Excel generation →
formula/link validation → model QC → [GATE 3: final QC] →
final Excel + build & diligence reports
```

- **Gate 1** — statements extracted, units/consolidation understood, historical
  periods identified, material errors flagged. Stops on material error (override
  with `--force`).
- **Gate 2** — an explicit internal `ModelSpec` (periods, currency, units,
  consolidation, statement architecture, schedules, valuation methods,
  exceptions).
- **Gate 3** — programmatic workbook QC: sheets exist, chronology oldest→newest,
  3-year forecast, formula integrity, cross-sheet reference validity, checks
  wired, no accidental forecast hardcodes, no broken references.

Intermediate artifacts are persisted (`data/raw`, `data/extracted`,
`data/normalized`, `data/validated`, `sources/source_registry`) so the
**RAW → NORMALIZED → VALIDATED → MODEL** chain is fully auditable.

---

## Architecture

```
financial-model-agent/
├── main.py                     # CLI entry point
├── config/                     # model / formatting / validation / accounting rules (YAML)
├── inputs/                     # documents + example company package
├── data/                       # raw → extracted → normalized → validated artifacts
├── sources/source_registry/    # persisted source database (provenance)
├── modules/
│   ├── schemas.py              # DataPoint / SourceDatabase / provenance (the backbone)
│   ├── config_loader.py        # config + accounting taxonomy
│   ├── pipeline.py             # end-to-end orchestration through the gates
│   ├── model_spec.py           # Gate 2 model specification
│   ├── ingestion/              # structured package loader
│   ├── document_analysis/      # PDF reading + document inventory
│   ├── extraction/             # PDF table → candidate datapoints (synonym-driven)
│   ├── normalization/          # units, sign, restatement engine
│   ├── disclosure_mapping/     # which schedules/sheets are relevant
│   └── validation/             # Gate 1 & Gate 2
├── excel/
│   ├── linking/cell_registry.py    # (sheet, metric, period) → address; lineage
│   ├── workbook_builder/           # two-pass builder + per-sheet modules
│   ├── formulas/                   # reusable formula/writer helpers
│   ├── formatting/                 # colour/number conventions
│   └── qc/                         # Gate 3 workbook QC
├── reports/                    # build report + diligence report generators
├── tests/                      # pytest suite + synthetic company fixture
└── outputs/models/             # generated workbooks + reports
```

### How linkage works

The workbook is built in **two passes**. Pass 1 (*layout*) reserves each sheet's
rows and registers the address of every `(sheet, metric, period)` cell in a
`CellRegistry`. Pass 2 (*fill*) writes the cells — and because every address is
already known, any sheet can emit a real cross-sheet reference
(`='Income Statement'!F16`) that is guaranteed to resolve. This is what makes the
model genuinely linked rather than a collection of duplicated hardcodes.

---

## Company-specific adaptation

The disclosure mapper decides which schedules to build from the *actual*
company's data — segments, operating drivers, working capital, capex, debt — so a
chemicals company, an IT-services company and a bank each get a workbook that
reflects their economics rather than a generic template.

---

## Testing

```bash
python -m pytest tests/ -q
```

The suite runs the full pipeline on a **synthetic test company** (5 historical
years, segments, working capital, capex, debt, a restated comparative and a
deliberately missing datapoint) and, when the optional `formulas` package is
installed, **numerically evaluates the generated workbook** to assert the balance
sheet balances and the cash flow ties in every period (historical *and* forecast).

---

## What the agent does **not** do

Forecast assumptions · business outlook · valuation assumptions (WACC, terminal
growth, multiples, peers, share count) · scenarios · investment thesis · target
price · conclusion. Those remain the analyst's intellectual work — the agent
eliminates the repetitive mechanical work while preserving analyst control.
