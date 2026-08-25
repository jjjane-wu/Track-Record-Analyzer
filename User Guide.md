# GP Track Record Analyzer — User Guide

*What the tool does, what you do, what you get.*

---

## Why this exists

Every GP submits its track record in a different Excel format — different
column names, layouts, units, footnotes. The analysis we want at the end is
the same every time: the standardized **Segmented Track Record Analysis
Output** workbook. This tool automates the part in between, so the hours you
would otherwise spend re-keying and reconciling a GP's spreadsheet go into
actually evaluating the manager.

> **Upload the GP's raw file → confirm a few things on screen → download the
> finished analysis workbook.** 3 steps, about a minute, same standard format
> every time.

Everything happens in your web browser — no spreadsheets to edit by hand, no
scripts to run. The app runs entirely on the local machine: GP data never
leaves it, and no external services are involved. (The optional last step —
publishing verified deals to the team database — copies data only into your
own team's SharePoint folder, nothing else.)

---

## Walkthrough — GP_2, from raw file to finished workbook

*(Follow along with `GP_2 Data Sheet 2025.09.30.xlsx`.)*

### Screen 1 — Upload & Parse

1. Open the app in your browser using the link you were given — it looks and
   works like an ordinary web page.
2. **Drag the GP's Excel file into the upload box** (or click to browse for
   it). The tool reads the file right away and shows:
   - **GP Name** — filled in automatically from the filename (for GP_2 it
     guesses `GP_2`). Correct it if the guess is wrong.
   - **Analyst Name** — type your name; it is recorded with the run.
   - **Workbook Structure** — a short summary of what the tool found in the
     file. For GP_2 it finds one consolidated deal table. If a GP splits
     deals across one tab per fund, you'll see the tabs it plans to combine,
     with controls to adjust if it guessed wrong. Usually you touch nothing
     here.
3. Click **Parse File →**. A few seconds later the app lists the funds it
   detected — for GP_2: Fund VIII, IX, X, XI, XII, VBP I — and **all of
   them are included automatically**. (To look at a subset later, use the
   Fund filter that every pivot in the output carries.)
4. Click **Next: Review Mapping →**.

### Screen 2 — Review Mapping  ← *the one step that needs your judgment*

The tool has matched the GP's raw columns to our 33 standard fields, and this
screen asks you to check its work. It is split into three colour-coded
blocks:

- **Needs Review (yellow)** — probable matches. Each row shows the standard
  field, the tool's reasoning ("↳ alias match: 'gross irr'"), a confidence
  score, and a **dropdown listing the GP's raw columns**. If a match is
  wrong, just open the dropdown and pick the right column. GP_2's
  `Gross IRR%(f) | 2025-09-30` header — a messy two-row header the tool
  stitched together — lands here and is correctly matched; one glance
  confirms it.
- **Unmapped fields (red, collapsed)** — fields the tool couldn't find.
  Expand the section and use the same dropdowns to map them yourself if the
  data exists in the file; otherwise they simply stay blank in the output.
- **Auto-confirmed (green, collapsed)** — the confident matches (Company,
  Fund, Sector, dates, invested capital …). Skim it once if you like.

Below the mappings sit **validation warnings** about the GP's own data (exit
date before entry date, MOIC ≤ 0, IRR that looks like percentage points).
These are flags on the *GP's* numbers, not the tool's — keep them for the
diligence call: inconsistencies in a manager's own track record file are
themselves diligence information.

Two minutes here is the whole job — a wrong mapping is the one mistake that
flows through everything. Click **Generate Analysis →**.

### Screen 3 — Download the Deal Level Input

The tool builds the hand-off file. What you see:

- **Summary metrics** — GP_2: 90 deals, 38 realized, 6 funds.
- **Unit conversion notice** — appears only when the tool had to rescale:
  GP_2 already reports in millions, so nothing shows; upload a GP reporting
  in thousands or absolute dollars and a banner lists exactly which columns
  were converted.
- **⬇️ Download Deal Level Input** — saves
  `[13-Jul-26 - GP_2] - Gross Deal Level Input.xlsx` (the date is the day
  you run it). A backup copy is kept automatically.
- **Step-by-step instructions** for the final step: open
  **TR-Analyzer.xlsm** — it comes with the app, in the **`vba`** folder of
  the folder you unzipped (click *Enable Macros* when Excel asks) — run the
  **`ImportInputsAndBuild`** macro (Opt+F8 / Alt+F8), pick the downloaded
  file — it imports the data and builds the 7 analysis tabs. Tip: save the
  built analysis under its own name (e.g. `GP_2 - Analysis.xlsm`) so the
  original analyzer stays empty and ready for the next GP.
- A **mapping log** and an **error log** at the bottom, for the record.

Something looks off? Click **← Back**, fix the mapping, download again,
re-run the macro — it takes seconds. For one-off data corrections (a value
the GP itself reported wrongly, a junk row that slipped through), edit the
**downloaded input file** in Excel and re-run the macro on it: that file is
the single source of truth. Never hand-patch the analysis workbook itself.

---

## Publishing to the database *(optional, after verification)*

Verified deals can go into the team's cross-GP **deal database** — the pool
that feeds the Power BI dashboard. This is deliberately a separate, manual
step: the app's automated output may still contain mapping or data errors,
so nothing enters the database until a person has checked it.

1. Open the corrected Deal Level Input file in Excel one last time — this
   exact file is what gets published.
2. In the app's sidebar, switch to **Publish to database**.
3. Upload the file. The app re-checks it and shows the results:
   - **Errors** (missing company names, text in number columns, unreadable
     dates) **block publishing** — fix them in Excel and upload again.
   - **Warnings** (blank funds, odd-looking IRRs, duplicate deals) don't
     block, but you must tick a box confirming you've checked them.
4. Click **Publish**. One CSV snapshot is written per GP per as-of date
   (`GP_2 - 2025-09-30.csv`) into the database folder, stamped with who
   published it and when.

Re-publishing the same GP and date **replaces** its snapshot — corrections
never create duplicates. A newer reporting date adds a new snapshot, so the
GP's history is kept. The page also lists everything currently in the
database.

The database folder is just a folder. Once it points at the team's
OneDrive-synced SharePoint folder (**Publish to database → Database
folder**), every publish uploads itself and the Power BI dashboard picks it
up on its next refresh — setup notes in `database/POWERBI_SETUP.md`.

---

## What you get — seven tabs, one line each

Every tab opens with a numbered, clickable contents list; the workbook opens
on a Table of Contents that jumps to any tab. The imported deal data lands
directly in the Deal List as plain values — the workbook carries no separate
inputs tab and no links back to the input file.

1. **Table of Contents** — navigation.
2. **Deal List** — the imported deal data (28 standard columns, incl. a
   per-fund currency; Realized Value shows an explicit 0 when the GP provided
   none) plus full per-deal analytics; the blue threshold tables set every
   bucket boundary (edit = instant sensitivity).
3. **Return & Loss Ratios** — pooled MOIC, Loss Ratio and Impaired
   Invested Capital across 15 cuts (sector, geography, vintage, fund, size,
   exit …), chart per cut. Loss Ratio = impaired value / invested capital
   (how much capital is currently lost); Impaired Invested Capital =
   capital sitting in below-1.0x deals / invested capital.
4. **Return Dispersion** — MOIC and IRR distributions: Count, % IC, and
   average per bucket — the outlier-dependence and left-tail view.
5. **Portfolio Construction** — capital mix by fund × sector/geography plus
   deal-count attributes — concentration, strategy drift, sourcing profile.
6. **Vintage Perf by Sector** — invested capital, MOIC and loss ratio by
   vintage (4 filters) plus vintage × sector count and pooled-MOIC matrices.
7. **Deployment & Exits** — capital deployment pacing (vintage × fund) and
   realization pacing (fund × exit year), pre-filtered to realized exits.

*(Four further tabs — Underperforming Assets, Partner Attribution, Op
Performance, Op Performance - Unrealized — are currently switched off at
Eric's request; they can be re-enabled later.)*

For metric definitions and how to read them, see **Metric Guide.md**.

---

## House rules

- **Blue = type here. White = calculated. Grey = key result.**
- **Every pivot has report filters** (Fund / Status / Hold-Period buckets;
  the operating tabs add Sector). Status = *Realized* is the acid test
  (GP_2: recalculates from 90 deals to 38) — the track record in cash terms,
  with the GP's own marks stripped out.
- **Charts never change** — they're fixed snapshots of the full portfolio
  (blank / n/a categories are left off the charts); the pivots are for
  exploring.
- **Deals missing a label are hidden from that pivot and its totals** — GP_2
  has 9 deals with no sector, so the Sector pivot totals 81, not 90. Nothing
  is deleted: un-hide via the pivot's filter dropdown. Missing MOIC/IRR shows
  as an "n/a" bucket.
- The filename really has square brackets — Excel's title bar just displays
  them as parentheses.
- Nothing leaves the machine; no external services.
