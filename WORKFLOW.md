# GP Track Record Parser — System Workflow

## Overview

The system takes a GP's raw track record Excel file (any format) and builds a
**Segmented Track Record Analysis Output** workbook — a clean, from-scratch
file with live pivot tables, charts, and report filters, following the format
of the IO reference templates. It runs as a Streamlit web app in the browser
at `http://localhost:8502`.

Pipeline in one line:

```
raw GP .xlsx  →  parse & map  →  review  →  [dd-mmm-yy - GP Name] - Segmented Track Record Analysis Output.xlsx
```

---

## How to Start the App

```bash
cd "Track Record Database"
./venv/bin/streamlit run app/app.py --server.port 8502
```

---

## User-Facing Flow (3 Screens)

### Screen 1 — Upload & Parse

The user uploads the GP's track record Excel file. The app immediately
profiles the workbook and shows:

- **GP Name** — auto-detected from the filename, editable
- **Analyst Name**
- **Workbook Structure** — the sheets/tables found, with the detected layout
  (consolidated vs per-fund tabs); the analyst can override which table is
  used or which tabs are combined
- **Select Funds to Include in Analysis** — after parsing, a fund multiselect
  (defaults to all detected funds)

Parsing runs the full pipeline (Stages 1–6 below) and moves to Screen 2.

### Screen 2 — Review Mapping

Displays the schema-inference result as a confidence-coded review table:

| Colour | Meaning |
|--------|---------|
| Green (auto-confirmed) | Confidence ≥ 0.85 — mapping is reliable |
| Yellow (needs review)  | 0.50 ≤ confidence < 0.85 — verify the suggested column |
| Red (unmapped)         | Confidence < 0.50 — field will be left blank |

The panel also shows **signal explanations** (why each mapping was chosen),
**validation issues** (data-quality warnings), and **unmapped source columns**.
The user can correct any mapping manually before proceeding — the corrected
mapping is what drives generation.

### Screen 3 — Generate & Download

- Transforms the selected funds' rows into standardised deal records
  (all deals are included — the excluded-deals logic is disabled by design)
- Shows a **unit normalisation** notice when monetary columns were rescaled
  to millions (e.g. a GP reporting absolute dollars or thousands)
- Builds the output workbook from a blank file and offers it for download as
  `[dd-mmm-yy - GP Name] - Segmented Track Record Analysis Output.xlsx`
  (date = processing day; the square brackets are real — Excel's title bar
  merely displays them as parentheses)
- Saves a copy to `outputs/` (the cross-GP CSV database append is
  currently disabled — parsed data isn't considered clean enough to
  accumulate yet; `csv_writer.py` remains available to re-enable later)
- Shows summary metrics, the full mapping log, and an error log if any
  phase failed

---

## Pipeline Stages (What Happens Under the Hood)

### Stage 1 — Workbook Profiling (`profiler.py`)

- Scans every worksheet for candidate tables, multi-row headers, merged cells,
  and orientation; scores candidates and picks the primary deal table
- Extracts workbook metadata: GP name candidates, report date, currency
- **Track Record Date detection** (`detect_track_record_date`): a labelled
  "as of" cell beats a date embedded in the filename (e.g. `12.31.2025`,
  `Q3 2025`) beats the profiler's first-seen date. This becomes the
  "Track Record Date" cell in the output; the *filename* uses today
- **Unit banner detection** (`detect_unit_banner`): scans banner rows for
  declarations like "($ in thousands)" / "in millions" — feeds Stage 4.5

### Stage 1.5 — Layout Detection (`profiler.group_candidate_tables`)

Decides how many tables feed the deal set:

- **Consolidated** — one sheet holds every deal (all 7 sample files)
- **Per-fund** — deals sharded across same-schema sibling tabs, detected by
  column-count proximity + header-token overlap and combined into one deal
  set; fund identity is recovered from tab names and rows are de-duplicated
  on `(company, fund, entry_date)`

No GP or sheet-name hardcoding; the UI shows the detected layout and lets the
analyst adjust it.

### Stage 2 — Table Extraction (`extractor.py`)

- Reads the chosen table(s), builds composite names from multi-row headers
- Profiles every column: data type, fill rate, header unit hints
  (`($M)`, `'000`, `(k)`, `%`, `x`), and distribution heuristics
  (`looks_like_moic`, `looks_like_irr`, `looks_like_currency`, …)

### Stage 3 — Schema Inference (`inferencer.py`)

Maps every raw column to one of the **33 standardised fields** using
independent signals; 1–3 combine by maximum, conflicts resolve greedily:

| Priority | Signal | Score range |
|----------|--------|-------------|
| 1 | Alias exact/substring match | 0.80 – 0.95 |
| 2 | Regex pattern match | 0.70 – 0.93 |
| 3 | Fuzzy token overlap | 0.55 – 0.85 |
| 4 | Column-profile boost | +0 – +0.20 |
| 5 | Semantic embedding (fallback below 0.50) | 0.50 – 0.80 |

**Company detection is value-based, not just header-based**: candidates whose
values look like URLs, long descriptions, numbers, or repetitive codes are
penalised, and when no header signals survive, a rescue scan finds the column
whose *values* look most like company names (fixes files where names sit
under junk headers like "Currency in $M" or in a "Position" column).

An optional LLM fallback interface exists but is disabled by default.

### Stage 3.5 — Junk-Row Filtering (`pipeline._drop_non_deal_rows`)

Drops non-deal rows before transformation: footnotes and annotations
(`* Please refer to…`, `- Investment valuations…`), fund-section banner rows
(fund names inside the company column with no dates), and rows with no deal
data across the signal columns.

### Stage 4 — Normalisation (`transformer.py`)

Per-row conversion into a standardised record:

- **Dates** parsed from many formats; **hold period** recomputed exactly from
  entry/exit dates
- **Status** normalised to exactly two states: `Realized` (incl. written-off)
  and `Unrealized` (incl. active / partially realized)
- **Fund vintage** inferred from each fund's earliest entry date
- Bucket labels (hold period, invested capital, MOIC, revenue, EBITDA, EV,
  entry multiple, margin, IRR) via shared threshold constants that match the
  output's editable helper tables

### Stage 4.5 — Monetary Unit Normalisation (`transformer.detect_monetary_scale`)

Every monetary figure in the output is in **millions**. Raw files that report
in absolute currency units (e.g. ILPA-format workbooks) or thousands are
rescaled at transform time. Signals, strongest first, each gated by value
plausibility so footnote letters like "EBITDA(k)" can't trigger them:

1. column-header unit hints (`'000` → ÷1,000; `($M)` → keep)
2. the file-level banner declaration from Stage 1
3. value magnitude with a whole-file consensus (a median that is impossible
   as millions decides the file's regime, keeping MOIC arithmetic consistent)

Screen 3 reports exactly which columns were rescaled.

### Stage 5 — Validation (`validator.py`)

GP-agnostic data-quality rules: field-level (MOIC > 0, IRR range, IC ≥ 0…),
row-level (exit ≥ entry, value vs invested sanity), table-level (required
fields mapped, IRR unit sanity). Severity: `error` or `warning`.

### Stage 6 — Review Report (`reviewer.py`)

Aggregates mappings + validation into the structured object that drives
Screen 2 (confirmed / review / unmapped fields, unmapped columns, findings).

---

## Output Builder (`build_output.py`)

The output workbook is built **from a blank file** — no heavy template is
round-tripped. openpyxl writes the data sheets; the pivot tables and charts
are injected as raw OOXML (openpyxl cannot create pivots). Structure — eight
tabs, in order (four further tabs are currently switched off — see below):

1. **Table of Contents** — numbered, banded list of internal hyperlinks to
   every other tab; the workbook opens here. New tabs appear automatically.

2. **Deal Level Inputs** — the clean input data as *values*: 28 columns
   (B..AC), meta block (GP Name / Track Record Date / Currency), table
   `GrossDealLevelInput`. Every data cell is a true input: light-blue fill,
   blue font. Per Eric's EWL revision: a **Fund Currency** column sits after
   Status (mapped from the raw file when a currency column exists — with a
   guard that rejects numeric values — else defaulted to the workbook
   currency), the Initial Invested Capital column is removed, and **Realized
   Value is written as an explicit 0** when the GP provided no value.

3. **Deal List** — the analysis table (`DealLevelInput`, 58 columns B..BG,
   header row 13, data from row 14) in the template's own formula language:
   input columns are blank-safe links to Deal Level Inputs, computed columns
   (Vintage, Hold Period, Total Value, Gross MOIC, all bucket columns incl.
   IRR Buckets) carry the template's structured formulas. Above the table:
   a tag row (Input / Formula / Entry / Exit), editable **bucket-threshold
   helper tables** (the only blue "input" cells on this tab), and deal-count
   helpers. Five key computed columns (Total Value, Gross MOIC, Performing,
   IC-in-loss, Impaired) are shaded light grey.

4. **Return & Loss Ratios** — 15 real pivots in template order (Sector,
   Geography, Process Type, GP Role, Exit Type, Revenue/EBITDA/Multiple/EV/
   Margin/IC buckets, Vintage ×2, COI, Fund), each with Count / MOIC /
   Loss-Ratio data fields (plus the Impaired and Capital-Deployment
   variants), Graph Label cells, and a combo chart (MOIC columns + Loss-Ratio
   line) cloned from the template's own chart XML.

5. **Return Dispersion** — Gross MOIC and Gross IRR dispersion sections:
   bucket pivots showing Count and **% IC** (share of invested capital,
   percent-of-column), with the template's colour-coded column charts
   (red → grey → greens). The n/a bucket stays in the pivot but not the chart.

6. **Portfolio Construction** — two Fund × dimension matrix pivots
   (Invested Capital by Fund and Sector / Geography, % of each fund's
   capital, percent-stacked column charts) and a "Deal Count Attributes"
   section: five Count-of-Company pivots with pie charts plus the total deal
   count. (The template's versions are data-model pivots; these are rebuilt
   as regular pivots with identical output.)

7. **Vintage Perf by Sector** — vintage pivot (Count / Invested Capital /
   MOIC / Loss Ratio, four report filters, combo chart) + three
   vintage × sector matrices (counts and pooled MOIC).

8. **Deployment & Exits** — four pivots, template-verbatim: InvCap % and
    Deal Count by vintage × fund (deployment pacing), then Exits % of IC by
    Fund and Exits by Year (fund × exit year, realization pacing) with a
    Status filter pre-selected to Realized. Rows/columns with no data under
    the default filters are omitted, matching what Excel shows on refresh.

**Switched-off tabs (EWL revision):** Underperforming Assets, Partner
Attribution, Op Performance and Op Performance - Unrealized are currently
excluded from the output at Eric's request ("don't need it yet"). Their
builder code is intact but commented out — search `build_output.py` for
`EWL` to re-enable (sheet creation + writer calls in `build_output()`, and
the matching blocks in `plan_extra()` / `_extra_jobs()`).

Key mechanics (all verified by opening real files in Excel):

- **One shared pivot cache** for all 33 pivots, shipped fully populated and
  kept **sheet-consistent**: cache and rendered values are recomputed with
  the sheet formulas' exact semantics, so the numbers do not change when
  Excel refreshes the pivots on open (current Excel honours refreshOnLoad).
- **Rendered cells**: every pivot's saved layout is written into the sheet so
  the tabs display instantly, and geometry survives refresh.
- **Report filters**: every pivot ships Fund / Status / Hold Period Buckets
  filters (a pivot whose own axis is one of these drops that filter), with
  head-room above each pivot so added filters don't overwrite headings.
- **Blank handling**: blank axis items are hidden via the pivot's native
  hidden-item filter — grand totals cover visible items only; the underlying
  deals remain and can be re-shown from the filter dropdown.
- **Charts carry literal data** (no cell references): each chart is a static
  snapshot of the full, unfiltered portfolio — filtering a pivot never
  changes a chart, and blank/n-a categories never appear in charts.
- **Styling convention**: blue = cell you type in; white = calculated;
  grey = key computed result. Tables carry no banding (template-faithful).

---

## Standardised Field Schema (33 fields)

Monetary fields are in millions of the deal currency (auto-normalised).

| Field ID | Label | Type |
|----------|-------|------|
| `fund` | Fund | string |
| `company` | Company | string |
| `region` | Geography | string |
| `sector` | Sector | string |
| `entry_date` | Entry Date | date |
| `exit_date` | Exit Date | date |
| `status` | Status | Realized / Unrealized |
| `role` | GP Role | string |
| `transaction_type` | Transaction Type | string |
| `competition` | Process Type | string |
| `sourcing_partner` | Sourcing Partner | string |
| `exit_type` | Exit Type | string |
| `holding_period` | Hold Period | decimal years |
| `ic_initial` | (no output column — kept as fallback for the IC bucket) | millions |
| `fund_currency` | Fund Currency (defaults to workbook currency; numeric values rejected) | string |
| `ic_total` | Total Invested Capital | millions |
| `realized` | Realized Value | millions |
| `unrealized` | Current Value | millions |
| `total_value` | Total Value | millions |
| `gross_moic` | Gross MOIC / TVPI | multiple (×) |
| `gross_irr` | Gross IRR | decimal |
| `entry_rev` | Entry LTM Revenue | millions |
| `entry_ebitda` | Entry LTM EBITDA | millions |
| `entry_net_debt` | Entry Net Debt | millions |
| `entry_ev` | Entry Enterprise Value | millions |
| `exit_rev` | Exit LTM Revenue | millions |
| `exit_ebitda` | Exit LTM EBITDA | millions |
| `exit_net_debt` | Exit Net Debt | millions |
| `exit_ev` | Exit Enterprise Value | millions |
| `valuation_method` | Valuation Method | string |
| `fund_ownership` | Fund Ownership % | percent |
| `no_of_seats` | Board Seats | integer |
| `coi_deal` | COI Deal | Yes / No / n/a |

---

## File & Module Map

```
Track Record Database/
├── app/
│   ├── app.py                        — Streamlit UI (3 screens)
│   ├── pipeline.py                   — Orchestrates Stages 1–6 + junk-row filtering
│   ├── profiler.py                   — Stage 1: structure, dates, unit banner
│   ├── extractor.py                  — Stage 2: extraction + column profiling
│   ├── inferencer.py                 — Stage 3: multi-signal schema inference
│   ├── transformer.py                — Stage 4/4.5: normalisation + unit scaling
│   ├── validator.py                  — Stage 5: data-quality rules
│   ├── reviewer.py                   — Stage 6: review report
│   ├── mapper.py                     — Field catalogue (TEMPLATE_FIELDS)
│   ├── parser.py                     — Low-level Excel reader
│   ├── build_output.py               — Output builder (sheets + pivot/chart OOXML)
│   ├── deal_list_spec.py             — Deal List schema: columns, formulas, helpers
│   ├── csv_writer.py                 — CSV database appender (currently not wired in)
│   ├── chart_template.xml            — Combo-chart blueprint (Return & Loss Ratios)
│   ├── chart_rd.xml                  — Dispersion-chart blueprint
│   ├── chart_pc_stacked.xml / _ser / _pie — Portfolio Construction chart blueprints
│   ├── chart_vintage.xml / chart_op.xml — Vintage & Op Performance chart blueprints
│   └── drawing_anchor_template.xml   — Chart anchor blueprint
├── outputs/                          — Auto-saved generated files (created at runtime; not in git)
├── database/                         — POWERBI_SETUP.md (deal-database appending disabled for now)
├── IO/                               — Reference input/output templates (not in git)
├── 1 - Example GP Track Records/     — Anonymised sample GP files (not in git)
├── venv/                             — Private Python environment (created by the launcher)
├── start.bat / start.command / start.sh — Double-click launchers (first run = setup)
├── requirements.txt                  — Core dependencies
├── README.md                         — Repo front page + analyst setup steps
├── User Guide.md                     — How to use the app, screen by screen
└── WORKFLOW.md                       — This document
```

**Dependency note:** the semantic column-matching fallback
(sentence-transformers / scikit-learn / torch — install into the venv
manually if wanted) is genuinely optional — `inferencer.py` wraps it in
try/except and degrades to the deterministic signals when the packages or
the model download are unavailable. Verified: the full pipeline parses the
samples on core-only dependencies.

---

## Design Principles

- **GP-agnostic**: no hardcoded column names, no GP-specific branching — every
  behaviour is a generalised algorithm (alias/regex/fuzzy/semantic matching,
  value-based detection, magnitude-based unit consensus).
- **Transparent**: every mapping carries a confidence score and explanation;
  unit rescaling and dropped rows are reported, never silent.
- **Template-faithful, built from blank**: the output reproduces the IO
  template's layout, formulas, styling, and chart designs without ever
  round-tripping the heavy template file.
- **Refresh-consistent**: what the file shows on open is exactly what Excel
  recomputes after a pivot refresh — values, layout, and hidden-blank state.
- **Charts are snapshots**: pivots are for interactive slicing; charts always
  show the full portfolio and are immune to filtering.
