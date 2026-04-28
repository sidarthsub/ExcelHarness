"""SQLite store of individual run results.

One row per (task × seed × iteration). Designed as a materialized view
over the `result.json` files under `benchmarks/runs/`, so it's always
rebuildable with `python -m benchmarks.experiments.store rebuild`.

Schema is intentionally minimal — enough to drive the leaderboard and
the Researcher's prompt context. Raw result.json stays on disk at the
path recorded in `result_json_path` for deep dives.
"""
from __future__ import annotations

import argparse
import json
import sqlite3
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

from benchmarks.experiments.loss import cold_cost_from_result, per_run_loss


BENCH_ROOT = Path(__file__).resolve().parent.parent
RUNS_ROOT = BENCH_ROOT / "runs"
DB_PATH = BENCH_ROOT / "experiments" / "runs.sqlite"


SCHEMA = """
CREATE TABLE IF NOT EXISTS runs (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    run_dir             TEXT NOT NULL UNIQUE,
    task_id             TEXT NOT NULL,
    tier                INTEGER,
    label               TEXT,
    seed                INTEGER,
    git_sha             TEXT,
    ts                  TEXT NOT NULL,

    accuracy            REAL,
    passed              INTEGER,
    total               INTEGER,
    completed           INTEGER,
    terminated_reason   TEXT,

    wall_seconds        REAL,
    time_budget_seconds REAL,
    cost_actual_usd     REAL,
    cost_cold_usd       REAL,
    cost_budget_usd     REAL,

    builder_model       TEXT,
    planner_used        INTEGER,

    -- Per-run diagnostics from harness.run_session (added with the
    -- harness migration). New runs populate these; legacy rows leave
    -- them NULL.
    oracle_seed_hits    INTEGER,
    oracle_cache_hits   INTEGER,
    oracle_llm_calls    INTEGER,
    oracle_over_budget  INTEGER,
    evaluator_calls     INTEGER,
    evaluator_cost_usd  REAL,
    fail_file_count     INTEGER,

    loss                REAL,
    loss_components_json TEXT,
    result_json_path    TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_runs_task_label ON runs(task_id, label);
CREATE INDEX IF NOT EXISTS idx_runs_ts ON runs(ts);
"""

# Idempotent migrations — applied on every connect() so older DBs pick up
# the new diagnostic columns without a rebuild. SQLite ignores ALTER TABLE
# ADD COLUMN if the column already exists *only* if we guard ourselves.
_MIGRATIONS = [
    "ALTER TABLE runs ADD COLUMN oracle_seed_hits INTEGER",
    "ALTER TABLE runs ADD COLUMN oracle_cache_hits INTEGER",
    "ALTER TABLE runs ADD COLUMN oracle_llm_calls INTEGER",
    "ALTER TABLE runs ADD COLUMN oracle_over_budget INTEGER",
    "ALTER TABLE runs ADD COLUMN evaluator_calls INTEGER",
    "ALTER TABLE runs ADD COLUMN evaluator_cost_usd REAL",
    "ALTER TABLE runs ADD COLUMN fail_file_count INTEGER",
]


def connect(db_path: Path = DB_PATH) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    for stmt in _MIGRATIONS:
        try:
            conn.execute(stmt)
        except sqlite3.OperationalError:
            # column already exists — that's fine, this is idempotent
            pass
    conn.commit()
    return conn


def _row_from_result(result: dict, run_dir: Path, label: str | None, seed: int | None,
                     git_sha: str | None) -> dict[str, Any]:
    components = per_run_loss(result)
    oracle_stats = result.get("oracle_stats") or {}
    evaluator_stats = result.get("evaluator_stats") or {}
    # Count fail-files left in the run dir. After the build loop the gate
    # forces the builder to delete each one as it's addressed; surviving
    # fail files mean either the builder gave up or the loop hit max_turns.
    try:
        fail_count = sum(1 for _ in run_dir.glob("eval_fail_*.json"))
    except Exception:
        fail_count = None
    return {
        "run_dir": str(run_dir),
        "task_id": result.get("task_id"),
        "tier": result.get("tier"),
        "label": label,
        "seed": seed,
        "git_sha": git_sha,
        "ts": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
        "accuracy": result.get("accuracy"),
        "passed": result.get("passed"),
        "total": result.get("total"),
        "completed": 1 if result.get("completed") else 0,
        "terminated_reason": result.get("terminated_reason"),
        "wall_seconds": result.get("wall_seconds"),
        "time_budget_seconds": result.get("time_budget_seconds"),
        "cost_actual_usd": result.get("dollars"),
        "cost_cold_usd": cold_cost_from_result(result),
        "cost_budget_usd": result.get("cost_budget_dollars"),
        "builder_model": result.get("builder_model"),
        "planner_used": 1 if result.get("spec_generated") else 0,
        # Oracle stats: planner P1 routes clarifications through
        # benchmarks.oracle.resolve_questions which returns a stats dict
        # with seed_hits / cache_hits / llm_calls / over_budget. Surfacing
        # these per-run lets the dashboard show how often the seed file
        # actually bites vs. how often we fall through to the LLM oracle.
        "oracle_seed_hits": oracle_stats.get("seed_hits"),
        "oracle_cache_hits": oracle_stats.get("cache_hits"),
        "oracle_llm_calls": oracle_stats.get("llm_calls"),
        "oracle_over_budget": oracle_stats.get("over_budget"),
        "evaluator_calls": evaluator_stats.get("calls"),
        "evaluator_cost_usd": evaluator_stats.get("cost_usd"),
        "fail_file_count": fail_count,
        "loss": components["loss"],
        "loss_components_json": json.dumps({k: v for k, v in components.items() if k != "loss"}),
        "result_json_path": str(run_dir / "result.json"),
    }


def insert_run(conn: sqlite3.Connection, result: dict, run_dir: Path,
               *, label: str | None = None, seed: int | None = None,
               git_sha: str | None = None) -> int:
    row = _row_from_result(result, run_dir, label, seed, git_sha)
    cols = ",".join(row.keys())
    placeholders = ",".join(["?"] * len(row))
    cur = conn.execute(
        f"INSERT OR REPLACE INTO runs ({cols}) VALUES ({placeholders})",
        tuple(row.values()),
    )
    conn.commit()
    return cur.lastrowid


def _discover_result_files() -> Iterable[Path]:
    if not RUNS_ROOT.exists():
        return []
    return sorted(RUNS_ROOT.glob("*/result.json"))


def rebuild(db_path: Path = DB_PATH) -> int:
    """Wipe and repopulate the DB from every result.json under runs/."""
    if db_path.exists():
        db_path.unlink()
    conn = connect(db_path)
    n = 0
    for p in _discover_result_files():
        try:
            result = json.loads(p.read_text())
        except Exception as e:
            print(f"skip {p}: {e}", file=sys.stderr)
            continue
        insert_run(conn, result, p.parent)
        n += 1
    conn.close()
    return n


def query_recent(conn: sqlite3.Connection, *, label: str | None = None,
                 limit: int = 50) -> list[sqlite3.Row]:
    if label is None:
        cur = conn.execute(
            "SELECT * FROM runs ORDER BY ts DESC LIMIT ?", (limit,)
        )
    else:
        cur = conn.execute(
            "SELECT * FROM runs WHERE label = ? ORDER BY ts DESC LIMIT ?",
            (label, limit),
        )
    return cur.fetchall()


def per_task_losses(conn: sqlite3.Connection, label: str) -> dict[str, dict]:
    """Return {task_id: {n, mean_loss, mean_accuracy, ...}} for one label.

    Used by the paired-by-task comparator in autoresearch — per-task
    means cancel most of the shared task-level variance between baseline
    and candidate runs, so deltas are far less noisy than comparing
    flat corpus loss.
    """
    cur = conn.execute(
        """SELECT task_id,
                  COUNT(*) n,
                  AVG(loss) mean_loss,
                  AVG(accuracy) mean_accuracy,
                  AVG(cost_cold_usd) mean_cost_cold_usd,
                  AVG(wall_seconds) mean_wall_seconds
             FROM runs
            WHERE label = ?
            GROUP BY task_id""",
        (label,),
    )
    return {r["task_id"]: dict(r) for r in cur.fetchall()}


def corpus_summary(conn: sqlite3.Connection, label: str) -> dict[str, Any]:
    """Aggregate metrics for a label — what the Researcher reads back."""
    cur = conn.execute(
        """SELECT COUNT(*) n, AVG(loss) loss, AVG(accuracy) acc,
                  AVG(cost_cold_usd) cost_cold, AVG(cost_actual_usd) cost_actual,
                  AVG(wall_seconds) wall,
                  SUM(CASE WHEN completed = 1 THEN 1 ELSE 0 END) * 1.0 / COUNT(*) completion_rate
             FROM runs WHERE label = ?""",
        (label,),
    )
    row = cur.fetchone()
    return dict(row) if row else {}


# ---- CLI ------------------------------------------------------------------


def _main() -> int:
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.store")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("rebuild", help="wipe the DB and re-ingest every result.json under runs/")
    q = sub.add_parser("recent", help="print the N most recent runs")
    q.add_argument("--label", default=None)
    q.add_argument("--limit", type=int, default=20)
    s = sub.add_parser("summary", help="aggregate stats for a label")
    s.add_argument("--label", required=True)
    args = ap.parse_args()

    if args.cmd == "rebuild":
        n = rebuild()
        print(f"rebuilt: {n} run(s) ingested")
        return 0
    conn = connect()
    if args.cmd == "recent":
        rows = query_recent(conn, label=args.label, limit=args.limit)
        for r in rows:
            print(json.dumps(dict(r), default=str))
        return 0
    if args.cmd == "summary":
        print(json.dumps(corpus_summary(conn, args.label), indent=2, default=str))
        return 0
    return 2


if __name__ == "__main__":
    sys.exit(_main())
