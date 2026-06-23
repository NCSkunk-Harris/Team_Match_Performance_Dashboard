#!/usr/bin/env python3
"""
Auto-rebuild watcher for the Courage Performance Dashboard.
============================================================

Watches the data spreadsheet, the template, and build_dashboard.py.
Whenever any of them is saved, it re-runs build_dashboard.py automatically,
so Courage_Team_Performance_Dashboard.html stays up to date without you
running any command.

HOW TO USE
----------
In Terminal:

    cd "/Users/tomharris/Desktop/Claude/Projects/Team Performance Dashboard"
    python3 "GitHub Automation/watch_and_build.py"

Leave it running in that window. Edit the Excel file (or the build script),
hit Save, and the dashboard rebuilds within a couple of seconds.
Press Ctrl+C to stop.

No extra packages required — uses only the Python standard library.
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# This script lives in "GitHub Automation/", so the project root is its parent.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_SCRIPT = PROJECT_ROOT / "build_dashboard.py"

# Files that should trigger a rebuild when changed.
WATCHED = [
    PROJECT_ROOT / "Data Source" / "NWSL Match Data - Team Level.xlsx",
    PROJECT_ROOT / "template_dashboard.html",
    BUILD_SCRIPT,
]

POLL_SECONDS = 2  # how often to check for changes


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def snapshot() -> dict:
    """Return {path: last-modified-time} for files that currently exist."""
    state = {}
    for f in WATCHED:
        try:
            state[f] = f.stat().st_mtime
        except FileNotFoundError:
            state[f] = None  # file missing right now (e.g. Excel mid-save)
    return state


def run_build() -> None:
    log("Change detected — rebuilding dashboard...")
    result = subprocess.run(
        [sys.executable, str(BUILD_SCRIPT)],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )
    if result.returncode == 0:
        log("Build complete ✅")
        # Surface any warnings the build printed.
        for line in result.stdout.splitlines():
            if "warn" in line.lower():
                log(f"  ⚠️  {line.strip()}")
    else:
        log("Build FAILED ❌ — details below:")
        print(result.stdout, flush=True)
        print(result.stderr, file=sys.stderr, flush=True)


def main() -> None:
    if not BUILD_SCRIPT.exists():
        log(f"Cannot find build script at {BUILD_SCRIPT}")
        sys.exit(1)

    log("Watching for changes. Press Ctrl+C to stop.")
    for f in WATCHED:
        log(f"  watching: {f.name}")

    last = snapshot()
    # Build once on startup so the HTML reflects current data immediately.
    run_build()

    try:
        while True:
            time.sleep(POLL_SECONDS)
            current = snapshot()
            # Trigger only on a real change to an existing file (ignore None,
            # which happens momentarily while Excel writes the file).
            changed = [
                f for f, t in current.items()
                if t is not None and t != last.get(f)
            ]
            if changed:
                # Wait briefly so a large save finishes before we read it.
                time.sleep(1)
                run_build()
                last = snapshot()
    except KeyboardInterrupt:
        log("Stopped. Goodbye!")


if __name__ == "__main__":
    main()
