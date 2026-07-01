[CmdletBinding()]
param(
    [string]$TaskName = "Family Finance Dashboard Auto Publish",
    [string]$DataPath = "",
    [string]$WorkbookPath = ""
)

$message = @"
Local finance dashboard auto-publish is retired.

The production dashboard now uses:
  - Google Sheets as the canonical source
  - Netlify Function: /api/finance-dashboard
  - Netlify Blobs as the latest-good snapshot cache
  - Netlify scheduled refresh every 15 minutes

No Windows scheduled task is required. Existing tasks named "$TaskName" can be disabled or deleted.
"@

Write-Host $message
