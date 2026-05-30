from __future__ import annotations

import argparse
import os
import sys
import traceback
from pathlib import Path

os.environ.setdefault("PYTHONDONTWRITEBYTECODE", "1")
sys.dont_write_bytecode = True

SCRIPT_DIR = Path(__file__).resolve().parent
LOG_PATH = SCRIPT_DIR / "local_autopublish.log"
sys.path.insert(0, str(SCRIPT_DIR))

from dashboard_publish import default_data_path, push_dashboard  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the family finance auto-publish without opening a console window.")
    parser.add_argument("--data", help="Absolute path to dashboard_data.json.")
    parser.add_argument("--workbook", help=argparse.SUPPRESS)
    return parser.parse_args()


def log(message: str) -> None:
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(message.rstrip() + "\n")


def resolve_data_path(candidate: str | None) -> Path:
    if candidate:
        preferred = Path(candidate).expanduser().resolve()
        if preferred.exists():
            return preferred
        log(f"Preferred dashboard_data.json missing, falling back: {preferred}")
    fallback = default_data_path()
    if fallback.exists():
        return fallback
    raise FileNotFoundError(f"No dashboard_data.json could be found for auto-publish: {fallback}")


def main() -> int:
    args = parse_args()
    if args.workbook:
        log(f"Ignoring retired workbook argument and watching dashboard_data.json directly: {args.workbook}")
    data_path = resolve_data_path(args.data)
    changed = push_dashboard(data_path=data_path)
    log("Published finance dashboard data update." if changed else "Checked dashboard_data.json; no publish needed.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        log(traceback.format_exc())
        raise SystemExit(1)
