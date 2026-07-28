# VBA Prototype — hybrid architecture (branch `vba`)

**Goal:** keep Python for what it's good at (parsing messy GP files, column
mapping with analyst review), move the *output building* into Excel VBA so
that (a) analysts only need Excel — no Python install — and (b) pivots are
**native**, built by Excel itself, eliminating the entire class of
hand-crafted-OOXML corruption bugs (removed PivotTable views, cache repairs,
255-char item limits, refresh drift) that the Python builder has to defend
against.

```
raw GP file ──► Python app (parse + map, unchanged)
                    │  produces the "Deal Level Inputs" sheet
                    ▼
        TR-Analyzer.xlsm  (this folder's modules imported once)
                    │  BuildAnalysisWorkbook()
                    ▼
        Deal List (74-col template formulas) + analysis tabs
        with REAL native pivots
```

## Files in this folder

| File | Role |
|---|---|
| `modSpec.bas` | **Generated — never edit by hand.** Deal List schema (74 columns, formulas, tag row), input column order, header block (bucket threshold tables), the 15 Return & Loss Ratios pivot specs. |
| `generate_vba_spec.py` | The generator for `modSpec.bas` — reads `app/deal_list_spec.py` + `app/build_output.py`, so the VBA and Python versions share one schema and cannot drift. Re-run after any spec change. |
| `modBuild.bas` | Builds the Deal List: meta, bucket helper tables, tag row, `DealLevelInput` table, all column formulas (same `in:/in0:/F:/FT:` notation as Python). |
| `modUtil.bas` | Shared plumbing: the one pivot cache, calculated fields, meta blocks, report filters, canonical bucket ordering, blank-hiding, mini contents lists. |
| `modCharts.bas` | Static snapshot charts: pivot values are copied via `GetPivotData` into a hidden `_ChartData` sheet right after each (still unfiltered) pivot is built, and plain charts are drawn over that block — so filtering a pivot never changes a chart, same rule as the Python build. |
| `modPivots.bas` | Return & Loss Ratios: 15 native pivots (Count / pooled MOIC / Loss Ratio, report filters, bucket order, hidden blanks) + a combo chart each. |
| `modDispersion.bas` | Return Dispersion: MOIC + IRR bucket pivots (Count / % IC / actual average) + % IC column charts (n/a bucket excluded from charts). |
| `modConstruction.bas` | Portfolio Construction: two Fund × Sector/Geography %-of-row matrices with stacked charts, five deal-count pivots with pies, total count. |
| `modVintage.bas` | Vintage Perf by Sector: the vintage performance pivot (4 filters, combo chart) + three Vintage × Sector matrices (counts and pooled MOIC). |
| `modDeployment.bas` | Deployment & Exits: InvCap % and Deal Count (vintage × fund), Exits % of IC and Exits by Year (fund × exit year, Status pre-set to Realized, n/a year hidden). |
| `modToc.bas` | Table of Contents: numbered internal links, moved to first position. |
| `modMain.bas` | Orchestration + entry points: `BuildAnalysisWorkbook()` (interactive) and `BuildHeadless()` (scripted). Each tab builds in its own error scope — one bad tab reports instead of killing the run — and sheets are arranged in the standard order at the end. |
| `build.vbs` | Windows COM driver: opens the template invisibly, injects the inputs sheet, runs the macro, saves a macro-free `.xlsx`, exits nonzero on failure. |
| `build.bat` | One-command wrapper for `build.vbs` (drag-and-drop friendly). |
| `mac_build.sh` | Mac dev helper: open + run macro + save via AppleScript. |
| `pipeline.bat` | The team's one command: raw GP file → `headless.py` parse → `build.vbs` → finished analysis (see above). |

The four `.bas` files are the version-controlled source; the assembled
`TR-Analyzer.xlsm` itself stays out of git (`*.xlsm` is ignored).

**All 8 output tabs are ported** (TOC, Deal Level Inputs consumed as input,
Deal List, Return & Loss Ratios, Return Dispersion, Portfolio Construction,
Vintage Perf by Sector, Deployment & Exits), including charts. Not ported:
pixel-exact template styling (fills, column widths, exact chart palettes) —
functional parity first; polish after the Windows validation run.

## Assembling the .xlsm (Windows, one-time)

1. Open Excel → blank workbook → save as `TR-Analyzer.xlsm`
   (macro-enabled). Keep it **outside** git — `*.xlsm` is ignored; the
   `.bas` files here are the version-controlled source.
2. `Alt+F11` → File → Import File… → import **all** `.bas` files in this
   folder (11 modules — `modSpec` must be among them).
3. Paste a Python-generated **Deal Level Inputs** sheet into the workbook
   (Move/Copy from any generated output, or paste values; keep the sheet
   name and the meta cells C3:C5).
4. `Alt+F8` → run `BuildAnalysisWorkbook`.

## The team pipeline — one command, raw file to finished analysis

Analysts never open the VBE: the module import below is a **one-time step
for whoever assembles the template**; the team receives the finished
`TR-Analyzer.xlsm`. Their whole workflow is then:

```bat
pipeline.bat "<raw GP track record>.xlsx"
```

(or drag the raw file onto `pipeline.bat`) — step 1 runs the Python parser
headless (`app/headless.py`: full pipeline, auto-accepted mapping, prints
the mapping summary), step 2 injects the inputs into the template and runs
the VBA build, saving `<raw> - Analysis.xlsx`. If step 1 prints
NEEDS-REVIEW / UNMAPPED fields, that GP deserves a pass through the
Streamlit app instead, where the mapping can be corrected by hand — the
pipeline is the happy path, the app is the judgment path.

## Scripted / batch runs (no clicking)

Yes — the whole build can run from one command. `build.vbs` drives Excel
through COM: opens the template invisibly, injects the "Deal Level Inputs"
sheet from any workbook, runs the macro, saves a plain `.xlsx` (macros
stripped), and exits with a proper error code.

```bat
build.bat "[23-Jul-26 - GP_2] - Segmented Track Record Analysis Output.xlsx"
```

…or drag an inputs workbook onto `build.bat`. Output lands next to the
input as `… - Analysis.xlsx`. On the Mac dev machine, `mac_build.sh`
does the open-run-save loop via AppleScript.

Two preconditions on Windows:

1. **Trusted Location** (one-time): add the folder holding
   `TR-Analyzer.xlsm` to File → Options → Trust Center → Trusted
   Locations — otherwise Excel silently disables macros for
   automation-opened files and the script fails.
2. `cscript.exe` must be allowed by corporate policy (AppLocker sometimes
   blocks it; the same COM calls can be ported to PowerShell if so).

## What to validate on a UNOIM laptop (the real questions)

1. **Does the macro run at all?** Files downloaded from the internet carry
   the Mark of the Web and Excel blocks their macros by default
   (2022 policy). Test: does right-click → Properties → *Unblock* appear,
   and does the macro run after unblocking? If corporate policy blocks
   macros entirely, this whole route is dead — better to know from a
   10-minute test than after porting everything.
2. Do the native pivots + calculated fields behave (filter to
   Status = Realized: MOIC re-pools correctly)?
3. Rough build time on a 200-deal file.

## Known gaps / notes

- Charts follow the "never change with filters" rule via the `_ChartData`
  snapshot mechanism (see `modCharts.bas`) — rebuilding refreshes them.
- Written blind on a Mac: expect the first Windows run to surface small
  object-model issues (chart constants, GetPivotData caption matching).
  Each tab is error-isolated, so the final dialog lists exactly which tab
  failed and why — send that text back for fixes.
- Mac Excel can import and *run* these modules, but validate on Windows —
  that's the deployment target.
