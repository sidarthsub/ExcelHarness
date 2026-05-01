"""Lightweight timing utility. Writes to a separate log so it can be added without restarting."""
import time
from pathlib import Path

_LOG = Path("/tmp/harness_timing.log")
_START = time.time()
_MARKS: dict[str, float] = {}


def mark(label: str) -> None:
    """Record a timestamp for a labeled event."""
    now = time.time()
    elapsed = now - _START
    _MARKS[label] = now
    line = f"[{elapsed:7.1f}s] {label}"
    print(line, flush=True)
    with _LOG.open("a") as f:
        f.write(line + "\n")


def since(label: str) -> float:
    """Seconds since a previous mark."""
    return time.time() - _MARKS.get(label, _START)


def reset():
    """Reset the start time and clear the log."""
    global _START
    _START = time.time()
    _MARKS.clear()
    _LOG.write_text("")
