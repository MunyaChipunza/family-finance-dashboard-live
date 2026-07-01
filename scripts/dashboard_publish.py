"""Deprecated local publish path.

The live dashboard no longer depends on local Excel, dashboard_data.json file
drops, a Windows scheduled task, or a Git push from this PC. Production data is
served by Netlify Functions from Google Sheets, with Netlify Blobs as the latest
successful snapshot cache.
"""

from __future__ import annotations

import argparse
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
DEFAULT_DATA = (BUNDLE_DIR / "dashboard_data.json").resolve()
DEPRECATION_MESSAGE = (
    "Local dashboard publishing is retired. The production dashboard reads "
    "/api/finance-dashboard, which refreshes from Google Sheets server-side."
)


def default_data_path() -> Path:
    return DEFAULT_DATA


def push_dashboard(*_args: object, **_kwargs: object) -> bool:
    print(DEPRECATION_MESSAGE)
    return False


def main() -> None:
    parser = argparse.ArgumentParser(description="Deprecated finance dashboard publisher.")
    parser.add_argument("--data", default=str(DEFAULT_DATA), help="Ignored. dashboard_data.json is not the production source.")
    parser.add_argument("--force", action="store_true", help="Ignored.")
    parser.add_argument("--commit-message", help="Ignored.")
    parser.parse_args()
    print(DEPRECATION_MESSAGE)


if __name__ == "__main__":
    main()
