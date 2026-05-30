from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import shutil
import subprocess
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
DEFAULT_DATA = (BUNDLE_DIR / "dashboard_data.json").resolve()
DEFAULT_STATE = SCRIPT_DIR / ".sync_state.json"
LOG_PATH = SCRIPT_DIR / "local_autopublish.log"
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
IGNORED_DIRTY_PATHS = {
    "dashboard_data.json",
    "scripts/.sync_state.json",
    "scripts/local_autopublish.log",
}
DEFAULT_COMMIT_PREFIX = "Update finance dashboard data"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Publish the finance dashboard when dashboard_data.json changes.")
    parser.add_argument("--data", default=str(DEFAULT_DATA), help="Path to dashboard_data.json.")
    parser.add_argument("--output", help=argparse.SUPPRESS)
    parser.add_argument("--workbook", help=argparse.SUPPRESS)
    parser.add_argument("--commit-message", help="Optional explicit Git commit message.")
    parser.add_argument("--force", action="store_true", help="Publish even if the JSON fingerprint has not changed.")
    return parser.parse_args()


def log(message: str) -> None:
    timestamp = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{timestamp} {message.rstrip()}\n")


def git_executable() -> str:
    path = shutil.which("git")
    if path:
        return path
    for candidate in (Path("C:/Program Files/Git/cmd/git.exe"), Path("C:/Program Files/Git/bin/git.exe")):
        if candidate.exists():
            return str(candidate)
    raise FileNotFoundError("Git executable not found.")


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    startupinfo = None
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
    return subprocess.run(
        [git_executable(), "-C", str(BUNDLE_DIR), *args],
        check=check,
        text=True,
        capture_output=True,
        creationflags=CREATE_NO_WINDOW,
        startupinfo=startupinfo,
    )


def default_data_path() -> Path:
    return DEFAULT_DATA


def find_default_workbook() -> None:
    """Retained for old runner imports; Excel is no longer a publish source."""
    return None


def resolve_data_path(candidate: str | Path | None) -> Path:
    path = Path(candidate).expanduser() if candidate else DEFAULT_DATA
    if not path.is_absolute():
        path = (BUNDLE_DIR / path).resolve()
    return path.resolve()


def is_git_repo() -> bool:
    result = run_git("rev-parse", "--is-inside-work-tree", check=False)
    return result.returncode == 0 and result.stdout.strip() == "true"


def has_origin() -> bool:
    result = run_git("remote", "get-url", "origin", check=False)
    return result.returncode == 0 and bool(result.stdout.strip())


def ensure_identity() -> None:
    name = run_git("config", "--get", "user.name", check=False)
    email = run_git("config", "--get", "user.email", check=False)
    if name.returncode == 0 and email.returncode == 0 and name.stdout.strip() and email.stdout.strip():
        return
    run_git("config", "user.name", "Family Finance Dashboard Sync")
    run_git("config", "user.email", "family-finance-dashboard-sync@local")


def sync_repo() -> None:
    fetch = run_git("fetch", "origin", "main", check=False)
    if fetch.returncode != 0:
        raise RuntimeError(fetch.stderr.strip() or "Could not fetch origin/main.")
    rebase = run_git("rebase", "origin/main", check=False)
    if rebase.returncode != 0:
        raise RuntimeError(rebase.stderr.strip() or rebase.stdout.strip() or "Could not rebase onto origin/main.")


def data_fingerprint(data_path: Path) -> dict[str, int | str]:
    stat = data_path.stat()
    return {
        "path": str(data_path),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def load_state() -> dict[str, object]:
    if not DEFAULT_STATE.exists():
        return {}
    try:
        return json.loads(DEFAULT_STATE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_state(payload: dict[str, object]) -> None:
    DEFAULT_STATE.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_dashboard_data(data_path: Path) -> dict[str, object]:
    try:
        with data_path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"dashboard_data.json is not valid JSON yet: {exc}") from exc
    except OSError as exc:
        raise RuntimeError(f"dashboard_data.json could not be read: {exc}") from exc

    if not isinstance(payload, dict):
        raise RuntimeError("dashboard_data.json must contain a JSON object.")
    return payload


def report_month(payload: dict[str, object]) -> str:
    value = payload.get("reportMonth")
    return str(value).strip() if value else "unknown month"


def commit_message_for(payload: dict[str, object], explicit: str | None = None) -> str:
    if explicit:
        return explicit
    return f"{DEFAULT_COMMIT_PREFIX} — {report_month(payload)}"


def needs_publish(data_path: Path, force: bool = False) -> bool:
    if force:
        return True
    current = data_fingerprint(data_path)
    state = load_state()
    return state.get("data") != current


def dirty_paths() -> list[str]:
    result = run_git("status", "--porcelain", check=False)
    paths: list[str] = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        path = line[3:].strip().replace("\\", "/")
        if path.startswith('"') and path.endswith('"'):
            path = path[1:-1]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        paths.append(path)
    return paths


def has_unrelated_changes() -> bool:
    for path in dirty_paths():
        if path not in IGNORED_DIRTY_PATHS:
            return True
    return False


def data_changed_in_git(data_path: Path) -> bool:
    rel = data_path.relative_to(BUNDLE_DIR).as_posix()
    status = run_git("status", "--short", "--", rel, check=False)
    return bool(status.stdout.strip())


def finalize_state(data_path: Path) -> None:
    save_state(
        {
            "data": data_fingerprint(data_path),
            "publishedAt": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
        }
    )


def push_dashboard(
    data_path: Path | str | None = None,
    commit_message: str | None = None,
    force: bool = False,
    **legacy_kwargs: object,
) -> bool:
    # Compatibility for the retired Excel runner signature:
    # push_dashboard(workbook_path=..., output_path=..., commit_message=...)
    if data_path is None:
        data_path = legacy_kwargs.get("output_path") or DEFAULT_DATA
    data_path = resolve_data_path(data_path)

    if not data_path.exists():
        raise FileNotFoundError(f"dashboard_data.json was not found: {data_path}")

    try:
        payload = load_dashboard_data(data_path)
    except RuntimeError as exc:
        message = f"Skipped auto-publish because {exc}"
        print(message)
        log(message)
        return False

    if not needs_publish(data_path, force=force):
        message = "No dashboard_data.json changes detected."
        print(message)
        log(message)
        return False

    message = commit_message_for(payload, explicit=commit_message)

    if not is_git_repo() or not has_origin():
        finalize_state(data_path)
        notice = "Detected dashboard_data.json changes, but no Git origin is configured; state recorded locally."
        print(notice)
        log(notice)
        return False

    if has_unrelated_changes():
        notice = "Skipped auto-publish because the dashboard repo has unrelated local changes."
        print(notice)
        log(notice)
        return False

    if not data_changed_in_git(data_path):
        finalize_state(data_path)
        notice = "dashboard_data.json fingerprint changed, but Git content is already up to date."
        print(notice)
        log(notice)
        return False

    ensure_identity()
    rel = data_path.relative_to(BUNDLE_DIR).as_posix()
    run_git("add", "--", rel)
    commit = run_git("commit", "-m", message, check=False)
    if commit.returncode != 0:
        notice = commit.stderr.strip() or commit.stdout.strip() or "Git commit failed."
        print(notice)
        log(f"Skipped auto-publish because {notice}")
        return False

    try:
        sync_repo()
    except RuntimeError as exc:
        notice = f"Committed dashboard_data.json locally, but skipped push because repo sync failed: {exc}"
        print(notice)
        log(notice)
        return False

    push = run_git("push", "-u", "origin", "main", check=False)
    if push.returncode != 0:
        notice = push.stderr.strip() or push.stdout.strip() or "Git push failed."
        print(notice)
        log(f"Committed dashboard_data.json locally, but push failed: {notice}")
        return False

    finalize_state(data_path)
    notice = f"Published dashboard_data.json with commit: {message}"
    print(notice)
    log(notice)
    return True


def main() -> None:
    args = parse_args()
    data_arg = args.data
    if args.output:
        data_arg = args.output
    if args.workbook:
        log(f"Ignoring retired workbook argument: {args.workbook}")
    push_dashboard(data_path=data_arg, commit_message=args.commit_message, force=args.force)


if __name__ == "__main__":
    main()
