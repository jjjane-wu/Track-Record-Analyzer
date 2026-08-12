# GP Track Record Analyzer — VBA hybrid edition (branch `vba`)

> **This branch** is the experimental hybrid version: Python keeps doing the
> parsing and column mapping, while the analysis workbook is built **inside
> Excel by VBA** with native pivot tables — so analysts need nothing but
> Excel, and the whole build can run from one command (`vba/build.bat`).
> Everything VBA-specific lives in the [`vba/`](vba/) folder — start with
> **[vba/README.md](vba/README.md)** for the architecture, assembly steps,
> and what to validate on a work laptop. The Python app below still works
> unchanged and remains the reference implementation.

Turns any GP's raw track record Excel file into the standardized
**Segmented Track Record Analysis Output** workbook — 8 tabs of live pivot
tables and charts — in about a minute, through a simple web page.

Upload the GP's file → confirm the column mapping on screen → download the
finished workbook. No spreadsheet re-keying, no scripts to run. Everything
runs locally on your own computer; GP data never leaves it.

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
workbook.

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
| `vba/` | **This branch's focus**: VBA output builder — importable modules, code generator, one-command build scripts (see `vba/README.md`) |
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
`.gitignore` (including a blanket ban on `*.xlsx`/`*.xls`/`*.pdf`). Before
any commit, run `git status` and confirm no data file is listed. Keep the
repository **private**.
