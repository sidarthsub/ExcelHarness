"""Structured shutdown for the autoresearch loop.

Usage:
    python -m benchmarks.experiments.shutdown

Reads the pidfile written by autoresearch at startup and kills the
entire process group, including orphaned Claude Agent SDK subprocesses,
eval_current subprocesses, and hidden Excel instances spawned by
xlwings. Idempotent — safe to run even if nothing is live.

Kill strategy (each step is best-effort):

  1. If the pidfile exists: SIGTERM the process group, wait ~2s, SIGKILL
     survivors. The autoresearch main process puts itself into its own
     process group at startup, so everything it spawned gets swept up.
  2. Fallback: match processes by command name
     (`benchmarks.experiments.autoresearch`, `eval_current`) and kill.
     Useful when the pidfile is missing (e.g., a past run died before
     writing it, or was started without the process-group hook).
  3. Kill any hidden Microsoft Excel processes. These are the
     `xw.App(visible=False)` instances; a user's foreground Excel would
     only be touched if they happen to have one open simultaneously
     (rare during an autoresearch run).
  4. Remove the pidfile.

Output is a human-readable summary of what was killed.
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import subprocess
import sys
import time
from pathlib import Path


BENCH_ROOT = Path(__file__).resolve().parent.parent
PIDFILE_PATH = BENCH_ROOT / "experiments" / "autoresearch.pid"
STATUS_PATH = BENCH_ROOT / "experiments" / "status.json"


def _read_pidfile() -> dict | None:
    if not PIDFILE_PATH.exists():
        return None
    try:
        return json.loads(PIDFILE_PATH.read_text())
    except Exception:
        return None


def _pg_alive(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
        return True
    except ProcessLookupError:
        return False
    except PermissionError:
        return True


def _kill_process_group(pgid: int, grace_seconds: float = 2.0) -> str:
    """SIGTERM the pgid; escalate to SIGKILL if anything survives."""
    try:
        os.killpg(pgid, signal.SIGTERM)
    except ProcessLookupError:
        return "no-such-pgid"
    # Wait for graceful teardown.
    for _ in range(int(grace_seconds * 10)):
        time.sleep(0.1)
        if not _pg_alive(pgid):
            return "SIGTERM ok"
    try:
        os.killpg(pgid, signal.SIGKILL)
        time.sleep(0.3)
        return "SIGKILL forced"
    except ProcessLookupError:
        return "died-during-wait"


def _pgrep(pattern: str, exact: bool = False) -> list[int]:
    """Wrap pgrep and return int PIDs. Returns [] if not found."""
    flag = "-x" if exact else "-f"
    try:
        out = subprocess.run(
            ["pgrep", flag, pattern],
            capture_output=True, text=True, check=False,
        )
    except FileNotFoundError:
        return []
    if out.returncode not in (0, 1):
        return []
    return [int(p) for p in out.stdout.split() if p.strip().isdigit()]


def _kill_pids(pids: list[int], label: str) -> list[int]:
    killed = []
    for pid in pids:
        try:
            os.kill(pid, signal.SIGKILL)
            killed.append(pid)
        except ProcessLookupError:
            pass
        except PermissionError:
            pass
    if killed:
        print(f"  killed {label}: {killed}")
    return killed


def shutdown() -> dict:
    summary: dict = {"pidfile_found": False, "pgid_kill": None,
                     "by_name_killed": [], "excel_killed": []}

    info = _read_pidfile()
    if info is not None:
        summary["pidfile_found"] = True
        pgid = info.get("pgid")
        print(f"pidfile found: pid={info.get('pid')} pgid={pgid} "
              f"started_at={info.get('started_at')}")
        if pgid is not None and pgid > 1:  # never kill init/session
            result = _kill_process_group(pgid)
            summary["pgid_kill"] = result
            print(f"  process-group kill: {result}")
    else:
        print("no pidfile — falling back to name-based cleanup")

    # Belt-and-suspenders: name-based cleanup for things that may have
    # detached from the process group (SDK binaries sometimes setsid).
    by_name = []
    for pat in (
        "benchmarks.experiments.autoresearch",
        "benchmarks.experiments.eval_current",
        "benchmarks.headless_builder",
    ):
        pids = _pgrep(pat, exact=False)
        by_name += _kill_pids(pids, f"by-name({pat})")
    summary["by_name_killed"] = by_name

    # Excel orphans. xlwings always spawns Microsoft Excel with its
    # exact process name, so this targets only the hidden instances.
    excel_pids = _pgrep("Microsoft Excel", exact=True)
    summary["excel_killed"] = _kill_pids(excel_pids, "Microsoft Excel")

    # Mark status and clean up.
    try:
        if STATUS_PATH.exists():
            existing = json.loads(STATUS_PATH.read_text())
            existing["phase"] = "stopped_externally"
            STATUS_PATH.write_text(json.dumps(existing, indent=2))
    except Exception:
        pass

    try:
        PIDFILE_PATH.unlink(missing_ok=True)
    except Exception:
        pass

    total_killed = (len(by_name) + len(summary["excel_killed"])
                    + (1 if summary["pgid_kill"] not in (None, "no-such-pgid") else 0))
    print(f"done. killed: process-group={summary['pgid_kill']!r}, "
          f"{len(by_name)} by-name, {len(summary['excel_killed'])} Excel instances.")
    if total_killed == 0:
        print("  (nothing was running)")
    return summary


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="benchmarks.experiments.shutdown",
        description="Kill the autoresearch loop cleanly (process group + Excel orphans).",
    )
    ap.parse_args()
    shutdown()
    return 0


if __name__ == "__main__":
    sys.exit(main())
