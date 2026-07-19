1. Open Power BI Desktop
2. Click Home → Get Data → Text/CSV
3. Navigate to: [this OneDrive folder] / database / gp_deals.csv
4. Select Import mode → click Load
5. Build three report pages:
   - Overview: KPI cards (total deals, avg MOIC, median hold period, top sector)
                Bar chart: avg MOIC by GP
                Scatter: MOIC vs hold period
   - Deal Browser: full table with slicers for GP, fund, sector, geography
                   Toggle slicer on "excluded" column
   - GP Comparison: side-by-side bar charts for selected GPs
6. To publish: Home → Publish → select your workspace
7. To schedule refresh: go to app.powerbi.com → dataset settings → Scheduled refresh
