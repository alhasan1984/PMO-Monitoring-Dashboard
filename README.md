# PMO Monitoring Dashboard

Oracle Customer Success Services executive dashboard for project revenue, unbilled position, milestones, hours KPIs, owner/portfolio lead visibility, and customer logo branding.

## Project Structure

```text
PMO-Monitoring-Dashboard/
  .github/workflows/      GitHub Pages deployment workflow
  docs/                   Project documentation
  scripts/                Local helper scripts
  src/                    Static web app source
    assets/               Oracle assets, customer logos, and leadership images
    app.js                Dashboard logic and workbook parser
    index.html            Main dashboard page
    styles.css            Dashboard styling
  package.json            Optional local commands
  README.md               Project overview
```

## Run Locally

From the repository root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\server.ps1
```

Then open:

```text
http://127.0.0.1:4173/
```

You can also open `src/index.html` directly in a browser.

## Current Scope

- Executive summary with fund, POC, revenue, hours, owner, portfolio lead, and customer logo.
- Quick revenue calculator on the main page.
- Revenue Estimator tab with quarterly and monthly drill-down charts and tables.
- Unbilled Situation tab with summary KPIs, comparison charts, and register.
- Milestones tab with status summaries, charts, and milestone register.
- Day and dark themes.
- Excel upload support for Master Sheet style workbooks.

## Workbook Notes

The dashboard expects project-level details on the `Master` sheet and joins related tabs by project name where available. Supported tabs include:

- `Master`
- `Projects TimeSheet`
- `Milestones`
- `PIC`

The customer logo can be stored in the `Customer Logo` column on the `Master` sheet as an Excel in-cell image.
