from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
from pathlib import Path

from dashboard_sync import DEFAULT_OUTPUT, DEFAULT_STATE, DEFAULT_WORKBOOK, refresh_dashboard_data


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
IGNORED_DIRTY_PATHS = {
    "dashboard_data.json",
    "scripts/.sync_state.json",
    "scripts/local_autopublish.log",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh the finance dashboard JSON and push when workbook data changed.")
    parser.add_argument("--workbook", help="Optional local workbook path.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output path for dashboard JSON.")
    parser.add_argument("--commit-message", default="Refresh finance dashboard data", help="Git commit message.")
    parser.add_argument("--force", action="store_true", help="Refresh even if the workbook fingerprint has not changed.")
    return parser.parse_args()


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


def find_default_workbook() -> Path | None:
    return DEFAULT_WORKBOOK if DEFAULT_WORKBOOK.exists() else None


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


def workbook_fingerprint(workbook_path: Path) -> dict[str, int | str]:
    stat = workbook_path.stat()
    return {
        "path": str(workbook_path),
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


def needs_refresh(workbook_path: Path, output_path: Path, force: bool = False) -> bool:
    if force or not output_path.exists():
        return True
    current = workbook_fingerprint(workbook_path)
    state = load_state()
    if state.get("workbook") != current:
        return True
    return output_path.stat().st_mtime_ns < workbook_path.stat().st_mtime_ns


def dirty_paths() -> list[str]:
    result = run_git("status", "--porcelain", check=False)
    paths: list[str] = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue
        path = line[3:].strip()
        path = path.replace("\\", "/")
        if path.startswith("\"") and path.endswith("\""):
            path = path[1:-1]
        paths.append(path)
    return paths


def has_unrelated_changes() -> bool:
    for path in dirty_paths():
        if path not in IGNORED_DIRTY_PATHS:
            return True
    return False


def output_changed(output_path: Path) -> bool:
    rel = output_path.relative_to(BUNDLE_DIR).as_posix()
    status = run_git("status", "--short", "--", rel, check=False)
    return bool(status.stdout.strip())


def finalize_state(workbook_path: Path, output_path: Path) -> None:
    save_state(
        {
            "workbook": workbook_fingerprint(workbook_path),
            "output": {
                "path": str(output_path),
                "mtime_ns": output_path.stat().st_mtime_ns if output_path.exists() else None,
            },
        }
    )


def push_dashboard(workbook_path: Path | None, output_path: Path, commit_message: str, force: bool = False) -> bool:
    if workbook_path is None:
        raise FileNotFoundError("No workbook was found for dashboard publishing.")

    if not needs_refresh(workbook_path, output_path, force=force):
        print("No workbook changes detected.")
        return False

    if not is_git_repo() or not has_origin():
        refresh_dashboard_data(workbook=str(workbook_path), output=output_path)
        finalize_state(workbook_path, output_path)
        print("Dashboard data refreshed locally.")
        return False

    if has_unrelated_changes():
        print("Skipped auto-publish because the dashboard repo has unrelated local changes.")
        return False

    sync_repo()
    refresh_dashboard_data(workbook=str(workbook_path), output=output_path)

    if not output_changed(output_path):
        finalize_state(workbook_path, output_path)
        print("Dashboard data is already up to date.")
        return False

    ensure_identity()
    rel = output_path.relative_to(BUNDLE_DIR).as_posix()
    run_git("add", "--", rel)
    run_git("commit", "-m", commit_message)
    run_git("push", "-u", "origin", "main")
    finalize_state(workbook_path, output_path)
    print("Dashboard data refreshed and pushed.")
    return True


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve() if args.workbook else find_default_workbook()
    output_path = Path(args.output).expanduser()
    if not output_path.is_absolute():
        output_path = (BUNDLE_DIR / output_path).resolve()
    push_dashboard(workbook_path=workbook_path, output_path=output_path, commit_message=args.commit_message, force=args.force)


if __name__ == "__main__":
    main()
