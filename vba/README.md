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

## What the prototype covers

| Module | Contents |
|---|---|
| `modSpec.bas` | **Generated** — Deal List schema (74 columns, formulas, tag row), input column order, header block (bucket threshold tables), the 15 Return & Loss Ratios pivot specs. Regenerate with `python3 vba/generate_vba_spec.py` after any spec change — never edit by hand. |
| `modBuild.bas` | Builds the Deal List: meta, bucket helper tables, tag row, `DealLevelInput` table, all column formulas (same `in:/in0:/F:/FT:` notation as Python). |
| `modPivots.bas` | Builds Return & Loss Ratios: 15 native pivots with Count / MOIC / Loss-Ratio (calculated fields → pooled, filter-correct math) + Fund/Status/HP-Buckets report filters. |
| `modMain.bas` | `BuildAnalysisWorkbook()` entry point. |

Not yet ported: Return Dispersion, Portfolio Construction, Vintage Perf by
Sector, Deployment & Exits, charts, TOC, exact template styling.

## Assembling the .xlsm (Windows, one-time)

1. Open Excel → blank workbook → save as `TR-Analyzer.xlsm`
   (macro-enabled). Keep it **outside** git — `*.xlsm` is ignored; the
   `.bas` files here are the version-controlled source.
2. `Alt+F11` → File → Import File… → import the four `.bas` files.
3. Paste a Python-generated **Deal Level Inputs** sheet into the workbook
   (Move/Copy from any generated output, or paste values; keep the sheet
   name and the meta cells C3:C5).
4. `Alt+F8` → run `BuildAnalysisWorkbook`.

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

- Calculated fields can't express *impaired-only* variants per pivot the way
  the Python cache does for every tab — fine for RLR, revisit per-tab.
- Native PivotCharts change when a pivot is filtered; the Python version
  ships static-snapshot charts by design. If the "charts never change" rule
  must hold, charts should be built from a hidden staging range instead.
- Mac Excel can import and *run* these modules, but validate on Windows —
  that's the deployment target.
