"""Evaluate the current working-tree harness on a task set and report corpus loss.

The Researcher calls this after every edit. It:

  1. Runs `harness.run_session` for each (task × seed) in the
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

import yaml

from benchmarks.experiments import store as store_mod
from benchmarks.experiments.loss import corpus_loss, per_run_loss, _tier_from_task_id
from benchmarks.oracle import OracleChannel
from harness import run_session


# Per-tier seed counts. t2 carries 2.5× corpus weight, so 1 t2 cell flipping
# accuracy 0.55→0.17 swings corpus_loss by ~0.2. Doubling t2 samples (3→4)
# shrinks t2-mean SE ~13% per tier-of-mean variance. t0 is near-saturated
# (acc≈1.0, σ≈0); 1 seed is sufficient. Total cells/iter: 9 → 7 — fewer
# wall-time minutes, better signal where it matters.
DEFAULT_SEEDS_PER_TIER: dict[int, int] = {0: 1, 1: 2, 2: 4}


def _seeds_for_task(task_id: str, seeds: int,
                    seeds_per_tier: dict[int, int] | None) -> int:
    if seeds_per_tier is None:
        return seeds
    tier = _tier_from_task_id(task_id)
    return seeds_per_tier.get(tier, seeds)


log = logging.getLogger("eval_current")


# ---- task sets -------------------------------------------------------------

# Canary retained for ad-hoc use only — the autoresearch loop no longer runs it.
# Visible drops t2 to keep iteration walls short (~10 min vs ~30 min). t2 signal
# comes via the holdout gate, which is the only place t2 tasks run during iteration.
CANARY_SET = ["t0_npv", "t1_inputs_from_term_sheet"]

# Default fixed sets — used when rotation isn't requested. 1 task per tier
# (t0+t1+t2) so the Researcher gets paired-Δ signal across the full
# difficulty curve. Visible t2 is `cap_table_series_ab` (mean ~300s/cell,
# baseline acc ~0.81 with real headroom) so per-iter wall stays ≤10 min.
# `lbo_mini` lives in holdout — its 600-1100s cell time and stronger
# failure-mode coverage make it the right promotion gate but the wrong
# fast-feedback signal.
# Visible/holdout swap (2026-05-04): the prior visible set
# (t0_npv / t1_inputs_from_term_sheet / t2_cap_table_series_ab) had t0
# saturated at 1.0 and t1 frozen — across 41 real iters, 0 t0 seed flips
# and 1 t1 seed flip out of 41×{1+2}=123 chances. All Researcher signal
# was concentrated in the 4 t2_cap_table seeds, where one seed flipping
# = 2.5 weighted units = the gate threshold itself. Net: pass-Δ noise
# floor equalled the accept threshold.
#
# Swap brings tasks with real headroom into visible:
#   - t1_revenue_build: build task with multi-stage failure modes
#   - t2_lbo_mini: prior holdout had mean acc ~0.47 (huge headroom)
# Holdout becomes the easier set — its job is detecting overfitting,
# not maximizing difficulty, so saturated holdout is fine.
VISIBLE_SET = [
    "t0_dcf_terminal_value",
    "t1_revenue_build",
    "t2_lbo_mini",
    "t2_cap_table_series_ab",  # 2nd t2 — diversifies the dominant signal source
]

HOLDOUT_SET = [
    "t0_npv",
    "t1_inputs_from_term_sheet",
    "t2_safe_convert_series_a",
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

# t2 visible pool — `cap_table_series_ab` is the fastest t2 (~300s/cell)
# with non-trivial headroom; chosen as the iter-time t2 representative.
# The slower/harder t2s stay in holdout so promotion gating still tests
# the full difficulty curve.
T2_VISIBLE_POOL = [
    "t2_cap_table_series_ab",
]

# Combined view kept for backward-compat callers that want the full set.
VISIBLE_POOL = T0_VISIBLE_POOL + T1_VISIBLE_POOL + T2_VISIBLE_POOL

T0_HOLDOUT_POOL = ["t0_dcf_terminal_value"]
T1_HOLDOUT_POOL = ["t1_revenue_build"]
T2_HOLDOUT_POOL = [
    "t2_lbo_mini",
    "t2_safe_convert_series_a",
]

# Combined view kept for backward-compat callers that want the full set.
HOLDOUT_POOL = T0_HOLDOUT_POOL + T1_HOLDOUT_POOL + T2_HOLDOUT_POOL


def rotate_sets(seed: int | None,
                n_visible_t0: int = 1, n_visible_t1: int = 1, n_visible_t2: int = 1,
                n_holdout_t0: int = 1, n_holdout_t1: int = 1,
                n_holdout_t2: int = 1) -> tuple[list[str], list[str]]:
    """Pick visible and holdout subsets, enforcing tier mix in both.

    Visible: 1 t0 + 1 t1 + 1 t2 — full difficulty range so the Researcher's
    paired-Δ signal includes pipeline failures (Planner/Oracle/Builder
    fragility) that only manifest on t2.
    Holdout: 1 t0 + 1 t1 + 1 t2 — guarantees disjoint coverage in the
    promotion-gate eval that the Researcher never directly sees.

    Deterministic given `seed`. Visible and holdout are disjoint — a task
    drawn for visible is excluded from the holdout pool.
    """
    import random
    rng = random.Random(seed if seed is not None else 0)
    visible = (
        rng.sample(T0_VISIBLE_POOL, n_visible_t0)
        + rng.sample(T1_VISIBLE_POOL, n_visible_t1)
        + rng.sample(T2_VISIBLE_POOL, n_visible_t2)
    )
    remaining_t0 = [t for t in T0_HOLDOUT_POOL if t not in visible]
    remaining_t1 = [t for t in T1_HOLDOUT_POOL if t not in visible]
    remaining_t2 = [t for t in T2_HOLDOUT_POOL if t not in visible]
    holdout = (
        rng.sample(remaining_t0, min(n_holdout_t0, len(remaining_t0)))
        + rng.sample(remaining_t1, min(n_holdout_t1, len(remaining_t1)))
        + rng.sample(remaining_t2, min(n_holdout_t2, len(remaining_t2)))
    )
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
    """Run one (task, seed) cell via harness.run_session.

    `skip_planner` is kept in the signature for back-compat with callers
    but is currently unused — the new harness always runs the planner.
    """
    log.info(f"running {cell.task_id} seed={cell.seed}")
    t0 = time.time()

    # Per-cell run_dir + ephemeral port for parallelism. Seed is mixed into
    # the port to avoid collision when multiple seeds of the same task run
    # concurrently. 3100-3999 range is unprivileged and unlikely to clash.
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
    run_dir = Path(__file__).resolve().parents[1] / "runs" / f"{ts}_{cell.task_id}_seed{cell.seed}"
    task_dir = Path(__file__).resolve().parents[1] / "tasks" / cell.task_id
    port = 3100 + (hash((cell.task_id, cell.seed)) % 800)

    # Adapt task.yaml → run_session kwargs (mirrors benchmarks/runner.cmd_run).
    # The harness no longer takes task_dir directly — eval substrate stages
    # brief / inputs / stub and plugs OracleChannel as the user simulator.
    task_meta = yaml.safe_load((task_dir / "task.yaml").read_text())
    brief = (task_dir / "brief.md").read_text() if (task_dir / "brief.md").exists() else ""
    inputs_dir = task_dir / "inputs" if (task_dir / "inputs").exists() else None
    stub_rel = task_meta.get("stub_file")
    stub_xlsx = (task_dir / stub_rel) if stub_rel else None
    if stub_xlsx and not stub_xlsx.exists():
        stub_xlsx = None
    grading_yaml = task_dir / "gold" / "grading.yaml"
    context_md = (task_dir / "context.md").read_text() if (task_dir / "context.md").exists() else ""
    seed_path = task_dir / "clarifications.seed.yaml"
    budget = task_meta.get("clarification_budget") or {}
    oracle_channel = OracleChannel(
        task_id=task_meta["id"],
        context_md=context_md,
        seed_path=seed_path if seed_path.exists() else None,
        max_questions=budget.get("max_questions", 10),
    )

    try:
        result = await run_session(
            brief=brief,
            inputs_dir=inputs_dir,
            run_dir=run_dir,
            user_channel=oracle_channel,
            stub_xlsx=stub_xlsx,
            output_keys=task_meta.get("output_keys") or [],
            grading_yaml=grading_yaml if grading_yaml.exists() else None,
            task_id=task_meta["id"],
            tier=task_meta.get("tier"),
            cost_budget_dollars=task_meta.get("cost_budget_dollars"),
            model=model,
            max_turns=max_turns,
            time_budget_seconds=time_budget,
            port=port,
        )
    except Exception as e:
        log.warning(f"{cell.task_id} seed={cell.seed} crashed: {type(e).__name__}: {e}")
        result = {
            "task_id": cell.task_id,
            "tier": None,
            "run_dir": str(run_dir),
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
    seeds_per_tier: dict[int, int] | None = None,
) -> dict[str, Any]:
    cells = [Cell(task_id=t, seed=s)
             for t in tasks
             for s in range(_seeds_for_task(t, seeds, seeds_per_tier))]

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
        "seeds_per_tier": seeds_per_tier,
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
                    help="Uniform seeds per task. Used as a fallback when "
                         "--seeds-per-tier is not set or doesn't list a tier.")
    ap.add_argument("--seeds-per-tier", default=None,
                    help="Per-tier seed counts as 'tier:n,...' (e.g. '0:1,1:2,2:4'). "
                         "Overrides --seeds for tiers it lists. Tiers come from task_id "
                         "prefix (t0/t1/t2). Pass 'default' to use DEFAULT_SEEDS_PER_TIER.")
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
    seeds_per_tier: dict[int, int] | None
    if args.seeds_per_tier is None:
        seeds_per_tier = None
    elif args.seeds_per_tier == "default":
        seeds_per_tier = dict(DEFAULT_SEEDS_PER_TIER)
    else:
        seeds_per_tier = {}
        for part in args.seeds_per_tier.split(","):
            tier_s, n_s = part.split(":")
            seeds_per_tier[int(tier_s.strip())] = int(n_s.strip())
    log.info(f"eval label={args.label} tasks={tasks} seeds={args.seeds} "
             f"seeds_per_tier={seeds_per_tier} model={args.model}")

    report = asyncio.run(run_eval(
        tasks=tasks,
        seeds=args.seeds,
        seeds_per_tier=seeds_per_tier,
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
