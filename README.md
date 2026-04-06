# Family Finance Live Dashboard

Live dashboard bundle for the family finance workbook.

## What is here

- `index.html` renders the finance dashboard.
- `dashboard_data.json` is the published data payload that the page reads.
- `scripts/refresh_dashboard_data.py` converts the workbook into `dashboard_data.json`.
- `scripts/publish_dashboard_data.py` refreshes the JSON and pushes it when the folder is inside a Git repo with an `origin`.
- `scripts/register_local_autopublish.ps1` creates a Windows scheduled task that republishes the dashboard every minute from this PC.

## Source workbook

The dashboard reads from:

- `../Finance_Input.xlsx`

## Typical local flow

1. Update `Finance_Input.xlsx`
2. Run `python scripts/refresh_dashboard_data.py`
3. Open `index.html` locally to preview
4. Run `python scripts/publish_dashboard_data.py` to refresh the JSON and push it

## Privacy note

This dashboard contains personal finance data. Treat the repo and any published URL as sensitive. If you decide to publish it publicly, first confirm exactly what data should remain visible.
