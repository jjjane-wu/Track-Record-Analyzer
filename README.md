# GP Track Record Analyzer — VBA hybrid edition (branch `vba`)

> **This branch** is the hybrid edition: the web app does the parsing and
> column mapping; the analysis itself is built **inside Excel by VBA** with
> native pivot tables. The analyzer workbook ships ready to use at
> [`vba/TR-Analyzer.xlsm`](vba/) (macros only — it contains no data), so
> analysts need nothing but Excel for the build step. Architecture, build
> scripts, and maintainer notes: **[vba/README.md](vba/README.md)**. The
> all-Python output builder remains in the code as the reference
> implementation.

Turns any GP's raw track record Excel file into the standardized
**Segmented Track Record Analysis** workbook — 7 tabs of native pivot
tables and charts — in about a minute.

Upload the GP's file → confirm the column mapping on screen → download the
standardized **Deal Level Input** → run one macro in the included
**TR-Analyzer.xlsm**. Verified deals can then be published to the
cross-GP deal database that feeds Power BI. Everything runs locally on
your own computer; GP data never leaves it.

---

## For analysts — getting started

You only do this once per computer:

1. **Install Python** (free): go to <https://www.python.org/downloads/> and
   install Python 3.10 or newer.
   *Windows:* on the first install screen, **tick "Add python.exe to PATH"**.
2. **Get this folder** onto your computer (ZIP from the maintainer or
   GitHub's *Code → Download ZIP*) and unzip it anywhere — Desktop, Documents
   and OneDrive folders are all fine.
3. **Double-click the launcher**:
   - Windows: `start.bat`
   - Mac: `start.command`

   The first run installs the app's components (needs internet, a few
   minutes). After that, the app starts in seconds and works offline.
4. Your browser opens the app automatically. Keep the launcher window open
   while you work; close it when you're done.

From there, follow **[User Guide.md](User%20Guide.md)** — the 3-screen
walkthrough of uploading a file, reviewing the mapping, and downloading the
Deal Level Input. The analysis itself is built by **`vba/TR-Analyzer.xlsm`**
(included in this folder): open it, click *Enable Macros*, and run
`ImportInputsAndBuild` on the downloaded file — the User Guide walks
through it.

**If something fails during setup:** screenshot the launcher window and send
it to the maintainer. The usual cause is an office network blocking Python
package downloads (IT can allow `pypi.org`).

---

## What's in this repository

| Path | What it is |
|------|------------|
| `app/` | The application — parsing pipeline, column mapper, and the output workbook builder |
| `start.bat` / `start.command` | Double-click launchers (Windows / Mac) — first run also installs everything |
| `requirements.txt` | Python components the app needs |
| `vba/` | **This branch's focus**: the VBA analyzer — `TR-Analyzer.xlsm` (ready to use, macros only), its version-controlled module source, code generator, and build scripts (see `vba/README.md`) |
| `User Guide.md` | How to use the app, screen by screen |
| `WORKFLOW.md` | Technical documentation of the pipeline (for maintainers) |
| `database/POWERBI_SETUP.md` | The cross-GP deal database (CSV snapshots → SharePoint → Power BI): how it works and how to connect |

Created at first run, not in the repository: `venv/` (the app's private
Python environment), `outputs/` (a backup copy of every generated workbook),
`database/deals/` (published database snapshots, until the database folder
is pointed at SharePoint).

---

## Data hygiene — read before committing

**This repository must never contain GP data.** All Excel/PDF files, raw GP
submissions, generated outputs, and reference templates are excluded by
`.gitignore` (including a blanket ban on `*.xlsx`/`*.xls`/`*.pdf`). The one
audited exception is `vba/TR-Analyzer.xlsm` — the empty analyzer (macros +
an instructions sheet, zero data). If `git status` ever shows it as
*modified*, a built copy has overwritten it — restore it, never commit it.
Before any commit, run `git status` and confirm no data file is listed.
Keep the repository **private**.
