# Family Finance Live Dashboard

Static dashboard UI plus Netlify server-side data bridge for the family finance command deck.

Live dashboard:

```text
https://munyachipunza.com/family-finance-dashboard-live/
```

## Production Architecture

`dashboard_data.json` is retained only as a local/static reference file. It is no longer the production source of truth.

Production flow:

1. The canonical finance source is a Google Sheet.
2. `index.html` fetches `/api/finance-dashboard` with `cache: "no-store"`.
3. The Netlify Function reads the Google Sheet server-side.
4. The function transforms the sheet rows into the existing dashboard JSON schema.
5. The latest successful payload is cached in Netlify Blobs.
6. If Google Sheets is unavailable, the API returns the latest Blob snapshot with `stale: true` and a `fallback` warning.
7. A scheduled Netlify Function refreshes the Blob snapshot every 15 minutes.

## Active Source Location

The healthy Git-connected local repo is:

```text
C:\Users\Dell\OneDrive\100. Zee\Finance\Finance_Dashboard_Live
```

There is an older Google Drive copy at:

```text
G:\My Drive\100. Zee\Finance\Finance_Dashboard_Live
```

That copy had an older README referring to `Finance_Input.xlsx`, while the previous Python scripts used `Finance.xlsx`. Treat both Excel references as retired. The active production source is Google Sheets.

## Google Sheet Shape

The Google Sheet should mirror the old single workbook tab:

```text
Budget!A:Z
```

The current migrated source Sheet is:

```text
https://docs.google.com/spreadsheets/d/1qSIt5r5MVWMWWKCqDifuV84-pB2ge5wCtsEOTYzvi0Y/edit
```

It was created from the current May 2026 dashboard payload and uses:

```text
Sheet1!A:Z
```

The transformer looks for a header row whose first four cells are:

```text
Status | Section | Group | Item
```

Supported standard columns:

```text
Status
Section
Group
Item
Owner
Budget Monthly
Actual This Month
Current Balance
Currency
Timing
Auto
Dashboard Tag
Priority
Notes
```

The report month is read from `B4` first, then `B5`, then falls back to the current month.

Optional upcoming payment rows can be added in the same sheet with:

```text
Section = Upcoming Payment
```

For those rows, the function reads:

```text
Item
Owner
Amount, Actual This Month, Budget Monthly, or Current Balance
Due Date, DueDate, or Timing
Days To Due, DaysToDue, or daysToDue
Auto
Category, Group, or Dashboard Tag
Priority
```

## Netlify Environment Variables

Set these in Netlify, not in source control:

```text
GOOGLE_SHEET_ID=...
GOOGLE_SHEET_RANGE=Budget!A:Z
GOOGLE_SERVICE_ACCOUNT_JSON=...
```

Alternatively, replace `GOOGLE_SERVICE_ACCOUNT_JSON` with:

```text
GOOGLE_SERVICE_ACCOUNT_EMAIL=...
GOOGLE_PRIVATE_KEY=...
```

If using a Google service account, share the Google Sheet with the service account email as a viewer. Otherwise the API will return a Google Sheets authorization error and the dashboard will fall back to the latest Blob snapshot.

The current deployment uses public-read Google Sheets CSV mode because no reusable Google service account credentials were available locally:

```text
GOOGLE_SHEET_ID=1qSIt5r5MVWMWWKCqDifuV84-pB2ge5wCtsEOTYzvi0Y
GOOGLE_SHEET_RANGE=Sheet1!A:Z
GOOGLE_SHEET_NAME=Sheet1
GOOGLE_SHEET_GID=0
GOOGLE_SHEETS_PUBLIC=true
```

In this mode the Sheet must have General access set to `Anyone with the link` / `Viewer`. To move back to private access later, unset `GOOGLE_SHEETS_PUBLIC`, add the service-account env vars above, and share the Sheet with the service account email.

Optional Notion snapshot:

```text
NOTION_TOKEN=...
NOTION_SNAPSHOT_PARENT_PAGE_ID=...
NOTION_SNAPSHOT_PAGE_ID=...
```

If `NOTION_TOKEN` and `NOTION_SNAPSHOT_PAGE_ID` are set, the scheduled refresh updates that page. If `NOTION_TOKEN` and `NOTION_SNAPSHOT_PARENT_PAGE_ID` are set but no page ID is set, the scheduled function creates a page called:

```text
Family Finance Live Dashboard - Current
```

The created page ID is stored in Netlify Blobs so future refreshes update the same page.

## Files

- `index.html` renders the dashboard and fetches `/api/finance-dashboard`.
- `netlify/functions/finance-dashboard.mts` serves the live JSON endpoint.
- `netlify/functions/finance-dashboard-refresh.mts` refreshes the Blob snapshot every 15 minutes.
- `netlify/functions/_shared/finance-transformer.ts` converts Google Sheet rows to the dashboard schema.
- `netlify/functions/_shared/google-sheets.ts` reads Google Sheets using service-account JWT auth.
- `netlify/functions/_shared/finance-cache.ts` stores and reads the latest good Netlify Blob snapshot.
- `netlify/functions/_shared/notion-snapshot.ts` optionally updates the Notion snapshot page.
- `scripts/*.py` and `scripts/*.ps1` are deprecated local automation stubs.

## Local Development

Install dependencies:

```text
npm install
```

Run type checks:

```text
npm run build
```

Run locally through Netlify:

```text
netlify dev
```

Then open:

```text
http://localhost:8888
http://localhost:8888/api/finance-dashboard
```

For local function testing, provide the Google env vars in a local `.env` file. Do not commit `.env`.

## Recovery

If the Google Sheet is down or credentials fail:

1. `/api/finance-dashboard` returns the latest successful Netlify Blob snapshot.
2. The payload includes `stale: true`.
3. `dataMode` becomes `Cached snapshot`.
4. `dataQuality` explains why the live refresh failed.
5. The UI continues loading instead of breaking.

If there is no Blob snapshot yet and Google Sheets also fails, the API returns HTTP `503`.

To recover:

1. Confirm the Netlify env vars are present.
2. Confirm the Google Sheet is shared with the service account email.
3. Open `/api/finance-dashboard` and check the JSON response.
4. Trigger or wait for the scheduled refresh.

## Retired Workflow

Do not use:

```text
Finance.xlsx
Finance_Input.xlsx
dashboard_data.json file drops
Windows Scheduled Task auto-publish
local git-push publisher
```

Those were the stale/brittle paths. They are no longer production infrastructure.

## Privacy

This dashboard contains personal finance data. Treat the repo, the Google Sheet, Netlify env vars, Blob snapshots, and any Notion snapshot page as sensitive.
