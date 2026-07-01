"""Deprecated local Excel sync path.

Production data now comes from the Netlify Function at /api/finance-dashboard,
which reads Google Sheets server-side and caches the latest successful payload
in Netlify Blobs. This module is kept only so old imports fail with a clear
message instead of silently regenerating stale dashboard_data.json from Excel.
"""

from __future__ import annotations

import argparse
from pathlib import Path


DEPRECATION_MESSAGE = (
    "The Excel-to-dashboard_data.json sync path is retired. "
    "Update the canonical Google Sheet and let /api/finance-dashboard refresh "
    "the dashboard through Netlify Functions and Netlify Blobs."
)


def refresh_dashboard_data(workbook: str | None = None, output: str | Path | None = None) -> Path:
    raise RuntimeError(DEPRECATION_MESSAGE)


def main() -> None:
    parser = argparse.ArgumentParser(description="Deprecated finance dashboard Excel sync.")
    parser.add_argument("--workbook", help="Ignored. Excel is no longer the dashboard source.")
    parser.add_argument("--output", help="Ignored. dashboard_data.json is no longer produced locally.")
    parser.parse_args()
    raise SystemExit(DEPRECATION_MESSAGE)


if __name__ == "__main__":
    main()
