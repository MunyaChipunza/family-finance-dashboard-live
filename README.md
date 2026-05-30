# Family Finance Live Dashboard

Static dashboard bundle for the family finance command deck.

## What is here

- `index.html` renders the live dashboard.
- `dashboard_data.json` is the single source of truth for the dashboard.
- `scripts/dashboard_publish.py` watches `dashboard_data.json`, commits changed data, and pushes to `origin main`.
- `scripts/run_local_autopublish.pyw` runs the publisher silently for the Windows scheduled task.
- `scripts/register_local_autopublish.ps1` creates or updates the Windows scheduled task.

## Publishing flow

1. Generate a ready `dashboard_data.json`.
2. Drop it into this folder, replacing the previous file.
3. The scheduled task checks every 5 minutes.
4. If the JSON file changed, the publisher commits it with the report month and pushes to GitHub.
5. Netlify deploys from `main`.

The retired Excel workbook is no longer used by the live publish workflow.

## Manual publish

Double-click:

```text
scripts/run_local_autopublish.pyw
```

Or run:

```text
python scripts/dashboard_publish.py
```

## Privacy note

This dashboard contains personal finance data. Treat the repo and published URL as sensitive.
