# Deal Database → Power BI

*How published snapshots become a live dashboard.*

---

## How the database works

The database is **one Excel workbook** — `TR Deal Database.xlsx` — holding
every GP's verified deals together, one row per deal, with the track record
date and GP name on every row. No server, no credentials.

- Every publish from the app (sidebar → **Publish to database**) merges that
  GP's rows into the workbook. Re-publishing the same GP + as-of date
  **replaces** those rows only — a correction never duplicates, and no other
  GP is touched. A new as-of date adds alongside, so the history of a GP's
  track record is preserved over time.
- **Every change is reversible.** Before each publish, remove or restore,
  the current workbook is copied into the `history/` subfolder; the app's
  Publish page can remove a GP snapshot or put the whole database back to
  any earlier state. An action log (who / what / when) rides along on a
  hidden `_Log` sheet inside the workbook.
- Only **verified** data enters: the analyst corrects and checks the Deal
  Level Input workbook in Excel first; the app validates it again (blocking
  on hard errors) before writing anything.

Power BI reads the one workbook directly.

---

## One-time setup — put the folder on SharePoint

Publishing "uploads" automatically without any code or APIs, because the
database folder *is* a synced SharePoint folder:

1. In the team's SharePoint site, create a folder in a document library,
   e.g. `Documents/TR Database`.
2. Open that folder in the browser and click **Sync** (or **Add shortcut to
   OneDrive**). It now appears in Finder/Explorer as a normal local folder.
3. In the app: **Publish to database → Database folder** → paste the local
   path of that synced folder → **Save folder**.

From then on, every publish updates `TR Deal Database.xlsx` in the folder
and the OneDrive client uploads it to SharePoint in the background. Until
SharePoint is set up, the default folder `database/deals/` next to the app
works fine — you can move the file and re-point the folder later. One
caveat of a single shared workbook: the app cannot write while someone has
the file open in Excel — it will say so and wait for you to close it.

---

## Connect Power BI

1. Power BI Desktop → **Get Data → Web** → paste the SharePoint link to
   `TR Deal Database.xlsx` (or **Get Data → Excel workbook** and browse to
   the synced local copy) → sign in.
2. Pick the **Deal Level Inputs** sheet (or the `GrossDealLevelInput`
   table) → **Transform Data**. One table, no combining step needed.
3. In Power Query, set column types once: `Track Record Date`, `Investment Date`,
   `Exit Date`, `Signing Date` → *Date* (values are ISO `YYYY-MM-DD`); monetary columns,
   `Gross TVPI`, `Gross IRR` → *Decimal Number*. Then **Close & Apply**.

**Latest snapshot per GP** (most dashboards want this): add a Power Query
step — *Group By* `GP` with aggregation *Max* of `Track Record Date`,
then merge that back and keep only matching rows. Keep the ungrouped table
too if you want "how did this GP's track record change over time" views.

## Publish & refresh

- **Home → Publish** to your workspace. Workspace access = who can see GP
  data; keep it restricted.
- app.powerbi.com → dataset → **Settings → Scheduled refresh**. Because the
  source is SharePoint Online (cloud-to-cloud), refresh works out of the box
  — **no gateway needed**. New publishes appear on the dashboard at the next
  refresh.
- If the folder is still local-only (no SharePoint yet), use **Get Data →
  Folder** instead; refresh then only happens in Desktop, or via a personal
  gateway on that machine.

## Suggested report pages

- **Overview** — KPI cards (GPs, deals, pooled MOIC, median hold period);
  bar: pooled MOIC by GP; scatter: MOIC vs hold period, colored by status.
- **Deal Browser** — full table with slicers for GP, fund, sector,
  geography, status.
- **GP Comparison** — side-by-side bars (pooled MOIC, loss ratio, DPI proxy
  = realized / invested) for selected GPs; deployment by vintage.
