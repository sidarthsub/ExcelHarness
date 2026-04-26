"""Evaluate the current working-tree harness on a task set and report corpus loss.

The Researcher calls this after every edit. It:

  1. Runs `headless_builder.run_headless` for each (task × seed) in the
     requested set. Runs are serial by default — the pseudo-bridge drives
     a hidden Excel instance via xlwings, which is not safe to spin up
     concurrently on macOS. `--parallel N` opts into concurrency once
     verified per-machine.
  2. Writes each result into the SQLite store tagged with `--label`.
  3. Computes and prints corpus loss + per-task breakdown as a single
     JSON object to stdout (the Researcher reads this).

Two standard task sets are wired in:

  - canary (--set canary)    : one task per tier, small. Fast signal.
  - visible (--set visible)  : the training split. Researcher sees these.
  - holdout (--set holdout)  : promotion gate. Researcher must NOT tune
                               against these directly.

You can also pass `--tasks t0_npv,t1_...` to override.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from benchmarks.experiments import store as store_mod
from benchmarks.experiments.loss import corpus_loss, per_run_loss
from benchmarks.headless_builder import run_headless


log = logging.getLogger("eval_current")


# ---- task sets -------------------------------------------------------------

# Canary retained for ad-hoc use only — the autoresearch loop no longer runs it.
# Visible drops t2 to keep iteration walls short (~10 min vs ~30 min). t2 signal
# comes via the holdout gate, which is the only place t2 tasks run during iteration.
CANARY_SET = ["t0_npv", "t1_inputs_from_term_sheet"]

# Default fixed sets — used when rotation isn't requested. Sized for fast
# iteration: 2 tasks × 2 seeds = 4 cells per visible eval (~5-7 min at parallel=3).
VISIBLE_SET = [
    "t0_npv",
    "t1_inputs_from_term_sheet",
]

HOLDOUT_SET = [
    "t0_dcf_terminal_value",
    "t2_lbo_mini",
]

# Full pools for rotation. Task tiers are roughly balanced for speed:
# t0s ~60s, t1s ~180s, t2s ~600s. Visible is split by tier and rotation
# enforces 1 t0 + 1 t1 — broader signal than picking any 2 from a flat
# pool (which routinely drew both t0s and missed multi-step build signal).
# Holdout includes t2s since it fires rarely.
T0_VISIBLE_POOL = [
    "t0_npv",
    "t0_option_pool_issuance",
    "t0_gross_up_post_money",
    "t0_pro_rata_allocation",
    "t0_runway_months",
    "t0_safe_conversion_price",
    "t0_waterfall_1x_nonpart",
    "t0_working_capital_delta",
    "t0ref_cap_table_fd_count",
    "t0ref_historical_cagr",
    "t0ref_term_sheet_to_price",
]

T1_VISIBLE_POOL = [
    "t1_inputs_from_term_sheet",
    "t1_debt_schedule",
    "t1_returns_table",
]

# Combined view kept for backward-compat callers that want the full set.
VISIBLE_POOL = T0_VISIBLE_POOL + T1_VISIBLE_POOL

HOLDOUT_POOL = [
    "t0_dcf_terminal_value",
    "t1_revenue_build",
    "t2_lbo_mini",
    "t2_cap_table_series_ab",
    "t2_safe_convert_series_a",
]


def rotate_sets(seed: int | None, n_visible_t0: int = 1, n_visible_t1: int = 1,
                n_holdout: int = 2) -> tuple[list[str], list[str]]:
    """Pick visible (1 t0 + 1 t1 by default) and holdout subsets from their pools.

    Deterministic given `seed`. Visible and holdout are disjoint — a task
    drawn for visible is excluded from the holdout pool so the Researcher
    can't accidentally tune against a task that also appears in its
    guardrail. Forcing one tier of each in visible prevents rotation luck
    from yielding all-t0 sessions (which give no multi-step build signal).
    """
    import random
    rng = random.Random(seed if seed is not None else 0)
    visible = rng.sample(T0_VISIBLE_POOL, n_visible_t0) + rng.sample(T1_VISIBLE_POOL, n_visible_t1)
    remaining_holdout = [t for t in HOLDOUT_POOL if t not in visible]
    holdout = rng.sample(remaining_holdout, min(n_holdout, len(remaining_holdout)))
    return visible, holdout


def resolve_task_set(name: str | None, explicit: str | None) -> list[str]:
    if explicit:
        return [t.strip() for t in explicit.split(",") if t.strip()]
    return {
        "canary":  CANARY_SET,
        "visible": VISIBLE_SET,
        "holdout": HOLDOUT_SET,
    }[name or "canary"]


# ---- runner ---------------------------------------------------------------


@dataclass
class Cell:
    task_id: str
    seed: int


def _git_sha() -> str | None:
    try:
        out = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=Path(__file__).resolve().parents[2],
            stderr=subprocess.DEVNULL,
        ).decode().strip()
        return out or None
    except Exception:
        return None


async def _one_run(cell: Cell, *, model: str, skip_planner: bool,
                   time_budget: float | None, max_turns: int) -> dict:
    log.info(f"running {cell.task_id} seed={cell.seed}")
    t0 = time.time()
    try:
        result = await run_headless(
            task_id=cell.task_id,
            model=model,
            time_budget_seconds=time_budget,
            max_turns=max_turns,
            skip_planner=skip_planner,
        )
    except Exception as e:
        # Synthesize a failure result so the corpus loss still accounts
        # for the cell rather than crashing the whole eval.
        log.warning(f"{cell.task_id} seed={cell.seed} crashed: {type(e).__name__}: {e}")
        result = {
            "task_id": cell.task_id,
            "tier": None,
            "run_dir": None,
            "candidate": None,
            "completed": False,
            "terminated_reason": f"crashed:{type(e).__name__}",
            "wall_seconds": time.time() - t0,
            "time_budget_seconds": time_budget or 600,
            "over_budget": False,
            "builder_model": model,
            "builder_usage": {},
            "planner_stats": {},
            "dollars": 0.0,
            "accuracy": 0.0,
            "passed": 0,
            "total": 0,
            "checks": [],
            "grading_error": str(e),
        }
    result["seed"] = cell.seed
    return result


async def run_eval(
    *,
    tasks: list[str],
    seeds: int,
    label: str,
    model: str,
    skip_planner: bool,
    time_budget: float | None,
    max_turns: int,
    parallel: int,
) -> dict[str, Any]:
    cells = [Cell(task_id=t, seed=s) for t in tasks for s in range(seeds)]

    results: list[dict] = []
    if parallel <= 1:
        for c in cells:
            results.append(await _one_run(
                c, model=model, skip_planner=skip_planner,
                time_budget=time_budget, max_turns=max_turns,
            ))
    else:
        sem = asyncio.Semaphore(parallel)

        async def _guarded(c: Cell):
            async with sem:
                return await _one_run(
                    c, model=model, skip_planner=skip_planner,
                    time_budget=time_budget, max_turns=max_turns,
                )

        results = list(await asyncio.gather(*[_guarded(c) for c in cells]))

    # Persist to store
    git_sha = _git_sha()
    conn = store_mod.connect()
    try:
        for r in results:
            run_dir = Path(r["run_dir"]) if r.get("run_dir") else Path(f"/tmp/crashed_{r['task_id']}_{r['seed']}")
            store_mod.insert_run(conn, r, run_dir, label=label, seed=r["seed"], git_sha=git_sha)
    finally:
        conn.close()

    # Per-task rollup
    by_task: dict[str, list[dict]] = {}
    for r in results:
        by_task.setdefault(r["task_id"], []).append(r)
    per_task: dict[str, dict] = {}
    for task_id, rs in by_task.items():
        agg = corpus_loss(rs)
        per_task[task_id] = {
            "n": agg["n"],
            "loss": round(agg["corpus_loss"], 4),
            "mean_accuracy": round(agg["mean_accuracy"], 4),
            "mean_cost_cold_usd": round(agg["mean_cost_cold_usd"], 4),
            "mean_cost_actual_usd": round(agg["mean_cost_actual_usd"], 4),
            "mean_wall_seconds": round(agg["mean_wall_seconds"], 1),
            "completion_rate": round(agg["completion_rate"], 3),
        }

    overall = corpus_loss(results)
    return {
        "label": label,
        "git_sha": git_sha,
        "tasks": tasks,
        "seeds": seeds,
        "n_runs": len(results),
        "corpus_loss": round(overall["corpus_loss"], 4),
        "mean_accuracy": round(overall["mean_accuracy"], 4),
        "mean_cost_cold_usd": round(overall["mean_cost_cold_usd"], 4),
        "mean_cost_actual_usd": round(overall["mean_cost_actual_usd"], 4),
        "mean_wall_seconds": round(overall["mean_wall_seconds"], 1),
        "completion_rate": round(overall["completion_rate"], 3),
        "eval_cost_usd": round(sum(r.get("dollars") or 0.0 for r in results), 4),
        "per_task": per_task,
    }


# ---- CLI ------------------------------------------------------------------


def _main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.eval_current")
    ap.add_argument("--set", dest="task_set", choices=["canary", "visible", "holdout"],
                    default="canary")
    ap.add_argument("--tasks", default=None,
                    help="Comma-separated task IDs. Overrides --set when provided.")
    ap.add_argument("--seeds", type=int, default=2,
                    help="Seeds per task. Default 2 — bump to 3 for promotion gates.")
    ap.add_argument("--label", required=True,
                    help="Identifier for this eval batch (researcher writes this).")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--skip-planner", action="store_true")
    ap.add_argument("--time-budget", type=float, default=900.0,
                    help="Global wall-time ceiling (seconds) per cell. "
                         "Default 900s (15 min). Task.yaml budgets are unchanged — "
                         "they still drive the time_loss normalization.")
    ap.add_argument("--max-turns", type=int, default=40)
    ap.add_argument("--parallel", type=int, default=2,
                    help="Concurrent runs. parallel=4 destabilizes hidden Excel under real workload "
                         "(AppleScript timeouts + workbook-not-found errors). 2 is the verified-safe ceiling "
                         "until we switch off ActiveWindow-based ops or batch xlwings commands.")
    ap.add_argument("--pretty", action="store_true")
    args = ap.parse_args()

    tasks = resolve_task_set(args.task_set, args.tasks)
    log.info(f"eval label={args.label} tasks={tasks} seeds={args.seeds} model={args.model}")

    report = asyncio.run(run_eval(
        tasks=tasks,
        seeds=args.seeds,
        label=args.label,
        model=args.model,
        skip_planner=args.skip_planner,
        time_budget=args.time_budget,
        max_turns=args.max_turns,
        parallel=args.parallel,
    ))
    print(json.dumps(report, indent=2 if args.pretty else None, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(_main())
