# Deal Database → Power BI

*How published snapshots become a live dashboard.*

---

## How the database works

The database is a **folder of CSV files** — no server, no credentials.

- Every publish from the app (sidebar → **Publish to database**) writes one
  snapshot: `GP_2 - 2025-09-30.csv` — that GP's verified deals as of that
  reporting date, one row per deal, with GP name, as-of date, and provenance
  (source file, who published, when) stamped on every row.
- Re-publishing the same GP + as-of date **replaces** its snapshot, so a
  correction never duplicates rows. A new as-of date adds a new snapshot, so
  the history of a GP's track record is preserved over time.
- Only **verified** data enters: the analyst corrects and checks the Deal
  Level Input workbook in Excel first; the app validates it again (blocking
  on hard errors) before writing anything.

Power BI reads the whole folder and combines the files into one long deal
table.

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

From then on, every published CSV lands in the folder and the OneDrive
client uploads it to SharePoint in the background. Until SharePoint is set
up, the default folder `database/deals/` next to the app works fine — you
can move the files and re-point the folder later.

---

## Connect Power BI

1. Power BI Desktop → **Get Data → SharePoint folder** → enter the *site*
   URL (e.g. `https://yourorg.sharepoint.com/sites/Investments`) → sign in.
2. Filter the file list to the `TR Database` folder → **Combine &
   Transform**. Power Query stacks every CSV into one table.
3. In Power Query, set column types once: `Track Record Date`, `Inv. Date`,
   `Exit Date` → *Date* (values are ISO `YYYY-MM-DD`); monetary columns,
   `Gross TVPI`, `Gross IRR` → *Decimal Number*. Then **Close & Apply**.

**Latest snapshot per GP** (most dashboards want this): add a Power Query
step — *Group By* `GP Name` with aggregation *Max* of `Track Record Date`,
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
