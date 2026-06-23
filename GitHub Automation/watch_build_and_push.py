#!/usr/bin/env python3
"""
Auto-rebuild + auto-push watcher for the Courage Performance Dashboard.
=======================================================================

Like watch_and_build.py, but after every SUCCESSFUL rebuild it also commits
the regenerated HTML and pushes to GitHub — so the live GitHub Pages site
updates without you running any git commands.

HOW TO USE
----------
In Terminal:

    cd "/Users/tomharris/Desktop/Claude/Projects/Team Performance Dashboard"
    python3 "GitHub Automation/watch_build_and_push.py"

Leave it running. Edit + save the Excel file (or the build script); the
dashboard rebuilds AND the change is pushed to GitHub automatically.
Press Ctrl+C to stop.

REQUIREMENTS
------------
- You must have already pushed once manually (so `git push` knows the remote)
  and saved your GitHub token, e.g. via:  git config --global credential.helper osxkeychain
  Otherwise the push step will hang waiting for a password.
- No extra Python packages required (standard library only).

SAFETY
------
- Pushes ONLY after a successful build (a failed build is never pushed).
- Skips the commit entirely if nothing actually changed.
- If git errors (lock file, auth, no network), it prints the error and keeps
  watching — it won't crash.
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BUILD_SCRIPT = PROJECT_ROOT / "build_dashboard.py"
OUTPUT_HTML = PROJECT_ROOT / "Courage_Team_Performance_Dashboard.html"

WATCHED = [
    PROJECT_ROOT / "Data Source" / "NWSL Match Data - Team Level.xlsx",
    PROJECT_ROOT / "template_dashboard.html",
    BUILD_SCRIPT,
]

POLL_SECONDS = 2
AUTO_PUSH = True  # set to False to rebuild locally without pushing


def log(msg: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def git(*args, check=False):
    """Run a git command in the project root. Returns the CompletedProcess."""
    return subprocess.run(
        ["git", *args],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
        check=check,
    )


def snapshot() -> dict:
    state = {}
    for f in WATCHED:
        try:
            state[f] = f.stat().st_mtime
        except FileNotFoundError:
            state[f] = None
    return state


def run_build() -> bool:
    """Run the build. Return True on success."""
    log("Change detected — rebuilding dashboard...")
    result = subprocess.run(
        [sys.executable, str(BUILD_SCRIPT)],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )
    if result.returncode == 0:
        log("Build complete ✅")
        for line in result.stdout.splitlines():
            if "warn" in line.lower():
                log(f"  ⚠️  {line.strip()}")
        return True
    log("Build FAILED ❌ — NOT pushing. Details:")
    print(result.stdout, flush=True)
    print(result.stderr, file=sys.stderr, flush=True)
    return False


def commit_and_push() -> None:
    if not AUTO_PUSH:
        return

    # Stage the regenerated dashboard (and any source edits you saved).
    git("add", "-A")

    # Anything staged? If not, skip — avoids empty commits.
    if git("diff", "--cached", "--quiet").returncode == 0:
        log("No changes to commit — skipping push.")
        return

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    commit = git("commit", "-m", f"Auto-update dashboard ({stamp})")
    if commit.returncode != 0:
        log("git commit failed:")
        print(commit.stderr or commit.stdout, file=sys.stderr, flush=True)
        return

    log("Committed. Pushing to GitHub...")
    push = git("push")
    if push.returncode == 0:
        log("Pushed to GitHub 🚀  (live site will update in ~1 min)")
    else:
        log("git push failed — your commit is saved locally, just not pushed.")
        print(push.stderr or push.stdout, file=sys.stderr, flush=True)
        log("Fix the issue (e.g. auth/network) and run `git push` manually.")


def preflight() -> bool:
    """Confirm we're in a git repo with a remote before starting."""
    if git("rev-parse", "--is-inside-work-tree").returncode != 0:
        log("This folder isn't a git repo yet. Run the GitHub setup first.")
        return False
    if AUTO_PUSH and not git("remote").stdout.strip():
        log("No git remote configured. Push once manually first (see SETUP_GITHUB.md).")
        return False
    return True


def main() -> None:
    if not BUILD_SCRIPT.exists():
        log(f"Cannot find build script at {BUILD_SCRIPT}")
        sys.exit(1)
    if not preflight():
        sys.exit(1)

    log("Watching for changes (auto-push ON). Press Ctrl+C to stop.")
    for f in WATCHED:
        log(f"  watching: {f.name}")

    last = snapshot()
    if run_build():
        commit_and_push()

    try:
        while True:
            time.sleep(POLL_SECONDS)
            current = snapshot()
            changed = [
                f for f, t in current.items()
                if t is not None and t != last.get(f)
            ]
            if changed:
                time.sleep(1)  # let the save finish
                if run_build():
                    commit_and_push()
                last = snapshot()
    except KeyboardInterrupt:
        log("Stopped. Goodbye!")


if __name__ == "__main__":
    main()
