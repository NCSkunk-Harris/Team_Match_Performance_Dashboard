#!/usr/bin/env python3
"""
On-command rebuild + push for the Courage Performance Dashboard.
================================================================

Runs ONCE when you call it (no background watching):
  1. Rebuilds the dashboard via build_dashboard.py
  2. If — and only if — the build succeeds, commits the changes
  3. Pushes to GitHub so the live page updates

HOW TO USE
----------
From Terminal:

    cd "/Users/tomharris/Desktop/Claude/Projects/Team Performance Dashboard"
    python3 "GitHub Automation/update_and_push.py"

OPTIONS
-------
    --no-push   Rebuild and commit locally, but do not push.
    --message "text"   Use a custom commit message.

SAFETY
------
- Pushes ONLY after a successful build (a failed build is never pushed).
- Skips the commit entirely if nothing actually changed.
- A non-zero exit code signals failure (useful if you wire this to a shortcut).
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_SCRIPT = PROJECT_ROOT / "build_dashboard_from_csv.py"  # CSV-first production build


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def git(*args):
    """Run a git command in the project root. Returns the CompletedProcess."""
    return subprocess.run(
        ["git", *args],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )


def run_build() -> bool:
    """Run the build. Return True on success."""
    log("Rebuilding dashboard...")
    result = subprocess.run(
        [sys.executable, str(BUILD_SCRIPT)],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )
    # Surface any warnings even on success so data gaps are visible.
    for line in result.stdout.splitlines():
        if "warn" in line.lower() or "⚠" in line:
            log(f"  {line.strip()}")
    if result.returncode == 0:
        log("Build complete.")
        return True
    log("Build FAILED — nothing will be committed or pushed. Details:")
    print(result.stdout, flush=True)
    print(result.stderr, file=sys.stderr, flush=True)
    return False


def preflight(do_push: bool) -> bool:
    """Confirm we're in a git repo (with a remote, if pushing)."""
    if git("rev-parse", "--is-inside-work-tree").returncode != 0:
        log("This folder isn't a git repo. Run the GitHub setup first (see SETUP_GITHUB.md).")
        return False
    if do_push and not git("remote").stdout.strip():
        log("No git remote configured. Push once manually first (see SETUP_GITHUB.md).")
        return False
    return True


def commit_and_push(do_push: bool, message: str | None) -> bool:
    git("add", "-A")

    # Anything staged? If not, skip — avoids empty commits.
    if git("diff", "--cached", "--quiet").returncode == 0:
        log("No changes to commit — dashboard is already up to date.")
        return True

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    msg = message or f"Update dashboard ({stamp})"
    commit = git("commit", "-m", msg)
    if commit.returncode != 0:
        log("git commit failed:")
        print(commit.stderr or commit.stdout, file=sys.stderr, flush=True)
        return False
    log(f"Committed: {msg}")

    if not do_push:
        log("Skipping push (--no-push). Commit is saved locally.")
        return True

    log("Pushing to GitHub...")
    push = git("push")
    if push.returncode == 0:
        log("Pushed. Live site will update in ~1 min.")
        return True
    log("git push failed — your commit is saved locally, just not pushed.")
    print(push.stderr or push.stdout, file=sys.stderr, flush=True)
    log("Fix the issue (e.g. auth/network) and run `git push` manually.")
    return False


def main() -> None:
    parser = argparse.ArgumentParser(description="Rebuild the dashboard and push, on command.")
    parser.add_argument("--no-push", action="store_true", help="Rebuild and commit locally without pushing.")
    parser.add_argument("--message", help="Custom commit message.")
    args = parser.parse_args()
    do_push = not args.no_push

    if not BUILD_SCRIPT.exists():
        log(f"Cannot find build script at {BUILD_SCRIPT}")
        sys.exit(1)
    if not preflight(do_push):
        sys.exit(1)

    if not run_build():
        sys.exit(1)
    if not commit_and_push(do_push, args.message):
        sys.exit(1)

    log("Done.")


if __name__ == "__main__":
    main()
