from __future__ import annotations

from datetime import datetime
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
LOG_PATH = SCRIPT_DIR / "local_autopublish.log"
MESSAGE = (
    "Local auto-publish is retired. The live dashboard now refreshes from "
    "Google Sheets through the Netlify Function at /api/finance-dashboard."
)


def log(message: str) -> None:
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{datetime.now().astimezone().isoformat(timespec='seconds')} {message.rstrip()}\n")


if __name__ == "__main__":
    log(MESSAGE)
