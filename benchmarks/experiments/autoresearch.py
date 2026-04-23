"""Autoresearch outer loop.

Spawns a Researcher agent (Claude Agent SDK) whose job is to iteratively
reduce corpus_loss by editing agents/*.md, benchmarks/headless_builder.py,
or benchmarks/pseudo_bridge.py.

Loop per iteration:

  1. Record the current baseline corpus_loss (ran once at the start).
  2. Ask the Researcher to produce ONE hypothesis, apply it, and run
     eval_current on the canary + visible sets. It writes a proposal
     file and prints its final JSON verdict.
  3. Read the latest eval from the store. If it strictly improves
     corpus_loss AND does not regress on the holdout set (we run that
     ourselves — Researcher can't), commit the change and update the
     baseline. Otherwise `git checkout .` to revert the working tree.
  4. Append a history row. Stop if max_iters or dollar budget hit.

The Researcher operates with a bounded tool allowlist that excludes
rubrics, gold models, and holdout tasks.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import logging
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from claude_agent_sdk import ClaudeSDKClient, ClaudeAgentOptions
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock

from benchmarks.experiments import store as store_mod
from benchmarks.experiments.eval_current import (
    HOLDOUT_SET, VISIBLE_SET, run_eval, resolve_task_set,
)


log = logging.getLogger("autoresearch")

ROOT = Path(__file__).resolve().parents[2]
AGENTS_DIR = ROOT / "agents"
BENCH_ROOT = ROOT / "benchmarks"
PROPOSALS_DIR = BENCH_ROOT / "experiments" / "proposals"
HISTORY_PATH = BENCH_ROOT / "experiments" / "history.jsonl"
LAST_EVAL_PATH = PROPOSALS_DIR / "last_eval.json"
STATUS_PATH = BENCH_ROOT / "experiments" / "status.json"

# Promotion gates.
CANARY_REGRESSION_TOL = 0.02   # Researcher bails out itself if canary worse than this.
IMPROVEMENT_EPSILON   = 0.005  # new loss must beat baseline by at least this much.
HOLDOUT_REGRESSION_MAX = 0.03  # holdout loss must not be more than this worse than pre-change holdout.

# Holdout is expensive (~20 min). Run it on accepted iters only, and only
# every Nth accepted iter — between firings the visible-set loss is the
# sole gate. Accumulated overfitting gets caught at the next holdout.
HOLDOUT_EVERY_N_ACCEPTED = 3

# Global wall-time ceiling per cell. Task.yaml budgets remain the intent
# (and drive time_loss normalization), but no cell is allowed to run past
# this — protects against hung Builders and runaway costs.
MAX_WALL_SECONDS = 900.0

# Default concurrency. 3 proved stable in testing; 4 destabilized under
# real workload before the ws.activate() + datetime JSON fixes landed.
DEFAULT_PARALLEL = 3


# ---- git helpers ----------------------------------------------------------


def _git(*args: str, check: bool = True) -> str:
    res = subprocess.run(
        ["git", *args], cwd=ROOT, check=check,
        capture_output=True, text=True,
    )
    return res.stdout.strip()


def _git_sha() -> str:
    return _git("rev-parse", "--short", "HEAD")


def _git_working_tree_clean() -> bool:
    out = _git("status", "--porcelain")
    return out == ""


def _revert_working_tree() -> None:
    # Restore every tracked file in the scope the Researcher is allowed to edit.
    for rel in (
        "agents/builder_v3.md",
        "agents/planner_v3.md",
        "benchmarks/headless_builder.py",
        "benchmarks/pseudo_bridge.py",
    ):
        p = ROOT / rel
        if p.exists():
            _git("checkout", "--", rel, check=False)


def _commit_iteration(label: str, loss: float) -> str:
    _git("add",
         "agents/builder_v3.md",
         "agents/planner_v3.md",
         "benchmarks/headless_builder.py",
         "benchmarks/pseudo_bridge.py",
         check=False)
    _git("commit", "-m",
         f"autoresearch {label}: corpus_loss -> {loss:.4f}",
         check=False)
    return _git_sha()


# ---- Researcher agent driver ---------------------------------------------


async def run_researcher_turn(iter_n: int, baseline: dict) -> str:
    """Spawn the Researcher agent for one hypothesis. Returns its final text."""
    system_prompt = (AGENTS_DIR / "researcher.md").read_text()

    # Files the Researcher can edit. Note: NO access to gold, grading.yaml,
    # task briefs, or the grader. Holdout tasks are path-excluded implicitly
    # because the Researcher never lists benchmarks/tasks/<holdout>/.
    allowed = [
        "Read", "Glob", "Grep",
        f"Edit({AGENTS_DIR}/builder_v3.md)",
        f"Edit({AGENTS_DIR}/planner_v3.md)",
        f"Edit({BENCH_ROOT}/headless_builder.py)",
        f"Edit({BENCH_ROOT}/pseudo_bridge.py)",
        f"Write({PROPOSALS_DIR}/iter_*.md)",
        # Running evals: scoped to the exact command surface.
        "Bash(python -m benchmarks.experiments.eval_current*)",
        "Bash(python3 -m benchmarks.experiments.eval_current*)",
        "Bash(cat benchmarks/experiments/proposals/*)",
        "Bash(cat benchmarks/experiments/history.jsonl)",
        "Bash(git diff*)",
        "Bash(git status*)",
    ]

    # Denylist — wrap with a reminder in the prompt since Agent SDK
    # allowlist already excludes everything not listed. Included for
    # clarity to the model.
    opts = ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=allowed,
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
        model="sonnet",
        max_thinking_tokens=5000,
        add_dirs=[str(ROOT / "agents"), str(ROOT / "benchmarks" / "experiments"),
                  str(ROOT / "benchmarks")],
    )

    # Seed the Researcher with the current state.
    history_tail = ""
    if HISTORY_PATH.exists():
        lines = HISTORY_PATH.read_text().splitlines()
        history_tail = "\n".join(lines[-20:])

    msg = (
        f"## Iteration {iter_n}\n\n"
        f"Current baseline corpus_loss: **{baseline['corpus_loss']:.4f}**\n"
        f"Per-task baseline:\n```json\n"
        f"{json.dumps(baseline.get('per_task', {}), indent=2)}\n```\n\n"
        f"Recent history (tail ~20):\n```\n{history_tail or '(empty)'}\n```\n\n"
        f"Proposal file: `{PROPOSALS_DIR}/iter_{iter_n}.md`\n"
        f"Labels: canary=`iter{iter_n}_canary`, visible=`iter{iter_n}_visible`\n\n"
        f"Do one iteration now: form a hypothesis, edit, run canary, "
        f"then (if non-regressive) run visible. Finish with your ACCEPT/REJECT "
        f"recommendation."
    )

    client = ClaudeSDKClient(options=opts)
    final_text = ""
    try:
        await client.connect()
        await client.query(msg)
        async for message in client.receive_messages():
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        final_text = block.text
            elif isinstance(message, ResultMessage):
                break
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass
    return final_text


# ---- promotion gate -------------------------------------------------------


def _load_latest_eval_for(label: str) -> dict | None:
    conn = store_mod.connect()
    try:
        summary = store_mod.corpus_summary(conn, label)
        return summary if summary and summary.get("n") else None
    finally:
        conn.close()


async def _run_holdout() -> dict[str, Any]:
    log.info("running holdout gate…")
    return await run_eval(
        tasks=HOLDOUT_SET,
        seeds=2,
        label=f"holdout_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}",
        model="sonnet",
        skip_planner=False,
        time_budget=MAX_WALL_SECONDS,
        max_turns=40,
        parallel=DEFAULT_PARALLEL,
    )


def _append_history(entry: dict) -> None:
    HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY_PATH.open("a") as f:
        f.write(json.dumps(entry, default=str) + "\n")


def _try_reuse_baseline(label_prefix: str, tasks: list[str],
                        min_seeds: int = 2) -> dict | None:
    """Reconstruct a baseline eval report from existing store rows.

    Returns None if we don't have enough coverage to avoid a fresh run.
    Requires at least `min_seeds` rows per task, all under a label starting
    with `label_prefix` (e.g. "baseline" or "holdout_"). Reads the raw
    result.json off disk to recompute corpus_loss via the current loss
    function — so any recent loss tweak is reflected in the reused number.
    """
    from benchmarks.experiments.loss import corpus_loss as _corpus_loss

    conn = store_mod.connect()
    try:
        rows: list[dict] = []
        for task_id in tasks:
            cur = conn.execute(
                """SELECT result_json_path FROM runs
                    WHERE label LIKE ? AND task_id = ?
                    ORDER BY id DESC LIMIT ?""",
                (f"{label_prefix}%", task_id, min_seeds),
            )
            cells = [dict(r) for r in cur.fetchall()]
            if len(cells) < min_seeds:
                return None
            rows.extend(cells)
    finally:
        conn.close()

    results: list[dict] = []
    for row in rows:
        p = Path(row["result_json_path"])
        if not p.exists():
            return None
        try:
            results.append(json.loads(p.read_text()))
        except Exception:
            return None

    agg = _corpus_loss(results)
    per_task: dict[str, dict] = {}
    for r in results:
        per_task.setdefault(r.get("task_id"), []).append(r)
    per_task_agg = {
        t: {
            "n": len(rs),
            "loss": round(_corpus_loss(rs)["corpus_loss"], 4),
            "mean_accuracy": round(_corpus_loss(rs)["mean_accuracy"], 4),
            "mean_cost_cold_usd": round(_corpus_loss(rs)["mean_cost_cold_usd"], 4),
            "mean_wall_seconds": round(_corpus_loss(rs)["mean_wall_seconds"], 1),
        }
        for t, rs in per_task.items()
    }
    return {
        "label": f"{label_prefix}_reused",
        "tasks": tasks,
        "seeds": min_seeds,
        "n_runs": len(results),
        "corpus_loss": round(agg["corpus_loss"], 4),
        "mean_accuracy": round(agg["mean_accuracy"], 4),
        "mean_cost_cold_usd": round(agg["mean_cost_cold_usd"], 4),
        "mean_cost_actual_usd": round(agg["mean_cost_actual_usd"], 4),
        "mean_wall_seconds": round(agg["mean_wall_seconds"], 1),
        "completion_rate": round(agg["completion_rate"], 3),
        "eval_cost_usd": 0.0,  # nothing spent — it's reused
        "per_task": per_task_agg,
        "reused": True,
    }


def _write_status(**fields: Any) -> None:
    """Persist a live status snapshot the dashboard can read.

    Fields are merged into the existing file so callers can update just
    one key (`phase`, `iter`, etc.) without clobbering the rest.
    """
    STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    existing: dict[str, Any] = {}
    if STATUS_PATH.exists():
        try:
            existing = json.loads(STATUS_PATH.read_text())
        except Exception:
            existing = {}
    existing.update(fields)
    existing["updated_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    STATUS_PATH.write_text(json.dumps(existing, indent=2, default=str))


# ---- outer loop -----------------------------------------------------------


async def outer_loop(*, max_iters: int, max_dollars: float,
                     skip_holdout: bool = False) -> None:
    PROPOSALS_DIR.mkdir(parents=True, exist_ok=True)

    if not _git_working_tree_clean():
        raise SystemExit(
            "refuse to start autoresearch with a dirty working tree. "
            "commit or stash first."
        )

    _write_status(phase="baseline", iter=0, started_at=datetime.now(timezone.utc).isoformat(),
                  max_iters=max_iters, max_dollars=max_dollars)

    # Baseline: try to reuse existing store rows before spending on a fresh eval.
    # min_seeds=1 is intentional — trading a bit of noise for ~15min of re-eval time.
    baseline_eval = _try_reuse_baseline("baseline", VISIBLE_SET, min_seeds=1)
    if baseline_eval is not None:
        log.info(f"reusing stored baseline — corpus_loss={baseline_eval['corpus_loss']:.4f} "
                 f"(n={baseline_eval['n_runs']} over {len(VISIBLE_SET)} tasks)")
    else:
        log.info("no usable stored baseline — running fresh on visible set…")
        baseline_eval = await run_eval(
            tasks=VISIBLE_SET, seeds=2, label="baseline",
            model="sonnet", skip_planner=False,
            time_budget=MAX_WALL_SECONDS, max_turns=40, parallel=DEFAULT_PARALLEL,
        )
    LAST_EVAL_PATH.write_text(json.dumps(baseline_eval, indent=2, default=str))
    baseline_loss = baseline_eval["corpus_loss"]
    log.info(f"baseline corpus_loss = {baseline_loss:.4f}")

    # Holdout baseline — only needed if we'll gate. Try reuse first.
    holdout_baseline = None
    if not skip_holdout:
        _write_status(phase="holdout_baseline")
        holdout_baseline = _try_reuse_baseline("holdout_", HOLDOUT_SET, min_seeds=1)
        if holdout_baseline is not None:
            log.info(f"reusing stored holdout baseline — "
                     f"corpus_loss={holdout_baseline['corpus_loss']:.4f}")
        else:
            log.info("no usable stored holdout — running fresh…")
            holdout_baseline = await _run_holdout()
        log.info(f"baseline holdout_loss = {holdout_baseline['corpus_loss']:.4f}")

    spent = (baseline_eval.get("eval_cost_usd") or 0.0) + \
            ((holdout_baseline.get("eval_cost_usd") if holdout_baseline else 0.0) or 0.0)
    current_sha = _git_sha()
    accepted_since_last_holdout = 0

    for i in range(1, max_iters + 1):
        if spent >= max_dollars:
            log.info(f"spend budget exhausted (${spent:.2f} ≥ ${max_dollars}); stopping.")
            break

        iter_start = time.time()
        log.info(f"=== iter {i} (spent=${spent:.2f}) ===")

        # Fresh branch for this attempt so revert is trivial.
        branch = f"autoresearch/iter_{i}_{datetime.now(timezone.utc).strftime('%H%M%S')}"
        _git("checkout", "-b", branch)

        _write_status(phase="researching", iter=i, branch=branch,
                      baseline_loss=baseline_loss, spent_usd=round(spent, 4),
                      iter_started_at=datetime.now(timezone.utc).isoformat())

        verdict_text = ""
        try:
            verdict_text = await run_researcher_turn(i, baseline_eval)
        except Exception as e:
            log.warning(f"researcher turn crashed: {type(e).__name__}: {e}")

        _write_status(phase="evaluating", iter=i)

        # Fresh eval is the Researcher's last `iter{i}_visible` batch.
        # Canary is no longer run — visible is small enough to be the primary gate.
        visible_label = f"iter{i}_visible"
        visible_summary = _load_latest_eval_for(visible_label)

        accepted = False
        reason = "no_eval"
        new_loss = None
        ran_holdout_this_iter = False

        if visible_summary and visible_summary.get("loss") is not None:
            new_loss = visible_summary["loss"]
            if new_loss + IMPROVEMENT_EPSILON >= baseline_loss:
                reason = f"no_improvement ({new_loss:.4f} vs {baseline_loss:.4f})"
            elif skip_holdout:
                accepted = True
                reason = "improved, holdout skipped"
            else:
                # Tentatively accept on visible; only fire holdout every Nth accept.
                would_be_accepted_count = accepted_since_last_holdout + 1
                if would_be_accepted_count >= HOLDOUT_EVERY_N_ACCEPTED:
                    _write_status(phase="holdout_gate", iter=i)
                    holdout_new = await _run_holdout()
                    ran_holdout_this_iter = True
                    spent += holdout_new.get("eval_cost_usd", 0.0) or 0.0
                    if (holdout_new["corpus_loss"]
                            > holdout_baseline["corpus_loss"] + HOLDOUT_REGRESSION_MAX):
                        reason = (f"holdout regression after {would_be_accepted_count} accepted — "
                                  f"{holdout_new['corpus_loss']:.4f} > "
                                  f"{holdout_baseline['corpus_loss']:.4f} + {HOLDOUT_REGRESSION_MAX}")
                    else:
                        accepted = True
                        reason = f"improved + holdout ok (after {would_be_accepted_count} accepted)"
                        holdout_baseline = holdout_new
                        accepted_since_last_holdout = 0
                else:
                    accepted = True
                    reason = (f"improved (visible-only; holdout due in "
                              f"{HOLDOUT_EVERY_N_ACCEPTED - would_be_accepted_count} more accepts)")
                    accepted_since_last_holdout = would_be_accepted_count
        else:
            reason = "researcher did not produce a visible eval"

        if accepted:
            sha = _commit_iteration(f"iter{i}", new_loss)
            baseline_loss = new_loss
            baseline_eval = {
                "corpus_loss": new_loss,
                "per_task": dict(visible_summary),
            }
            LAST_EVAL_PATH.write_text(json.dumps(baseline_eval, indent=2, default=str))
            _git("checkout", "main", check=False)
            _git("merge", "--ff-only", branch, check=False)
            current_sha = sha
        else:
            _git("checkout", "--", ".", check=False)
            _git("clean", "-fd", "benchmarks/experiments/proposals", check=False)
            _git("checkout", "main", check=False)
            _git("branch", "-D", branch, check=False)

        # Track costs from this iter's evals.
        spent += (visible_summary or {}).get("cost_actual", 0.0) or 0.0
        spent += (canary_summary or {}).get("cost_actual", 0.0) or 0.0

        _append_history({
            "iter": i,
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "accepted": accepted,
            "reason": reason,
            "baseline_loss": baseline_loss,
            "new_loss": new_loss,
            "sha_before": current_sha,
            "wall_seconds": round(time.time() - iter_start, 1),
            "spent_usd": round(spent, 4),
            "ran_holdout": ran_holdout_this_iter,
            "accepted_since_last_holdout": accepted_since_last_holdout,
            "researcher_final_text": verdict_text[-800:],
        })
        _write_status(phase="idle_between_iters", iter=i,
                      baseline_loss=baseline_loss, spent_usd=round(spent, 4))
        log.info(f"iter {i}: accepted={accepted} reason={reason!r} loss={new_loss}")

    _write_status(phase="done", baseline_loss=baseline_loss, spent_usd=round(spent, 4))
    log.info(f"done. spent=${spent:.2f}, baseline_loss={baseline_loss:.4f}")


# ---- CLI -----------------------------------------------------------------


def _main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.autoresearch")
    ap.add_argument("--max-iters", type=int, default=10)
    ap.add_argument("--max-dollars", type=float, default=50.0)
    ap.add_argument("--skip-holdout", action="store_true",
                    help="Skip the holdout gate. Only for debugging the loop itself.")
    args = ap.parse_args()
    asyncio.run(outer_loop(
        max_iters=args.max_iters,
        max_dollars=args.max_dollars,
        skip_holdout=args.skip_holdout,
    ))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
