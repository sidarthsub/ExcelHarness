"""Autoresearch outer loop.

Spawns a Researcher agent (Claude Agent SDK) whose job is to iteratively
reduce corpus_loss by editing agents/*.md, harness.py, bridge.py, pseudo_bridge.py,
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
import os
import re
import signal
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from claude_agent_sdk import ClaudeSDKClient, ClaudeAgentOptions
from claude_agent_sdk.types import AssistantMessage, ResultMessage, TextBlock

from benchmarks.experiments import store as store_mod
from benchmarks.experiments.eval_current import (
    DEFAULT_SEEDS_PER_TIER, HOLDOUT_SET, VISIBLE_SET, run_eval,
    resolve_task_set, rotate_sets,
)
from benchmarks.experiments.loss import _tier_from_task_id


log = logging.getLogger("autoresearch")

ROOT = Path(__file__).resolve().parents[2]
AGENTS_DIR = ROOT / "agents"
BENCH_ROOT = ROOT / "benchmarks"
PROPOSALS_DIR = BENCH_ROOT / "experiments" / "proposals"
HISTORY_PATH = BENCH_ROOT / "experiments" / "history.jsonl"
LAST_EVAL_PATH = PROPOSALS_DIR / "last_eval.json"
STATUS_PATH = BENCH_ROOT / "experiments" / "status.json"
PIDFILE_PATH = BENCH_ROOT / "experiments" / "autoresearch.pid"

# Promotion gates.
CANARY_REGRESSION_TOL = 0.02   # Researcher bails out itself if canary worse than this.
IMPROVEMENT_EPSILON   = 0.015  # paired delta must be at least this negative (continuous gate, kept for reporting).

# Pass-rate gate — primary accept criterion. The continuous loss-delta gate
# was noise-dominated: with 3-seed multi-modal accuracy distributions
# (e.g. {0.11, 0.33, 1.0}) the per-seed swings dwarfed the 0.015 threshold,
# so noise rejects looked identical to real rejects in history.jsonl and
# the Researcher couldn't learn from them.
#
# Pass-rate uses the discrete order statistic instead: count #seeds per task
# with accuracy >= PASS_ACCURACY_THRESHOLD, weighted by tier. A real fix
# (e.g. resolving the Y1 timing convention) flips one or more seeds from
# fail → pass, which moves pass-count by ≥1 — cleanly above the noise floor.
# Mean-aggregate, by contrast, would average that flip with surviving
# failure modes and produce a much smaller continuous delta.
PASS_ACCURACY_THRESHOLD = 0.9
MIN_PASS_COUNT_DELTA_WEIGHTED = 1.0  # weighted Σ(w_t × Δpass_t); tier weights are 1.0/1.5/2.5
MAX_PER_TASK_REGRESSION = 1          # reject if any task lost ≥(this+1) seeds, even with a net win

# Hybrid gate — secondary path that catches improvements pass-rate misses:
# pure cost/speed wins (no accuracy change), and sub-threshold accuracy
# gains that move continuous loss substantially without flipping seeds
# across the 0.9 pass-line. Calibrated to ~3× the observed noise floor
# (~0.10 spurious deltas in the prior 26-iter session) so it filters
# noise while accepting real meaningful continuous improvements.
CONTINUOUS_BIG_WIN_THRESHOLD = 0.10  # corpus_delta must be ≤ -this for secondary path

HOLDOUT_REGRESSION_MAX = 0.03  # holdout loss must not be more than this worse than pre-change holdout.

# Holdout is expensive (~20 min). Run it on accepted iters only, and only
# every Nth accepted iter — between firings the visible-set loss is the
# sole gate. Accumulated overfitting gets caught at the next holdout.
HOLDOUT_EVERY_N_ACCEPTED = 3

# Global wall-time ceiling per cell. Task.yaml budgets remain the intent
# (and drive time_loss normalization), but no cell is allowed to run past
# this — protects against hung Builders and runaway costs.
# Per-cell wall caps. Visible runs use the tighter budget so iter wall
# stays ≤ 10 min; holdout keeps headroom for the harder lbo_mini task
# whose cells legitimately take 6-12 min.
VISIBLE_TIME_BUDGET = 800.0
HOLDOUT_TIME_BUDGET = 1100.0
MAX_WALL_SECONDS = HOLDOUT_TIME_BUDGET  # back-compat alias

# Default concurrency. 3 proved stable in testing; 4 destabilized under
# real workload before the ws.activate() + datetime JSON fixes landed.
DEFAULT_PARALLEL = 9


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


# Patterns that flag a tier-leak: the change introduces a branch keyed on
# benchmark-only metadata that won't exist in production sessions. Linted
# against the iter's diff (added lines only) BEFORE running the visible
# eval, so we don't waste eval cost on a change that's already disqualified.
_TIER_LEAK_PATTERNS = [
    # Lookups into task_meta for benchmark-only fields.
    re.compile(r'task_meta\s*\.\s*get\s*\(\s*["\'](?:tier|cost_budget_dollars|time_budget_seconds|output_keys|stub_file)["\']'),
    re.compile(r'task_meta\s*\[\s*["\'](?:tier|cost_budget_dollars|time_budget_seconds|output_keys|stub_file)["\']'),
    # Direct task.yaml reads.
    re.compile(r'task\.yaml\b'),
    re.compile(r'yaml\.safe_load\([^)]*task\b'),
]


def _check_for_tier_leaks(branch: str) -> str | None:
    """Scan added lines in this iter's diff for benchmark-only metadata refs.

    Returns a reject reason string, or None if the diff is clean. Only
    inspects files in the Researcher's editable scope so legitimate reads
    of task_meta inside files we own (e.g. harness.run_session reads tier
    just for the result.json metadata) don't false-positive — those lines
    aren't ADDED in this iter's diff.

    Comment-aware: for .py files, content past a `#` is stripped before
    matching, so a Researcher comment like
    "# Production-safe: not task.yaml metadata" doesn't false-positive.
    """
    scope = [
        "agents/builder_v3.md",
        "agents/planner_v3.md",
        "agents/evaluator_v3.md",
        "harness.py",
        "bridge.py",
        "pseudo_bridge.py",
    ]
    diff = _git("diff", "main", "--unified=0", "--", *scope, check=False)
    current_file: str | None = None
    for line in diff.splitlines():
        # Track which file's diff we're in so we can strip Python comments.
        if line.startswith("+++ "):
            # `+++ b/path/to/file`
            parts = line.split(maxsplit=1)
            current_file = parts[1].lstrip("b/") if len(parts) > 1 else None
            continue
        if line.startswith("--- "):
            continue
        # Only inspect ADDED lines.
        if not line.startswith("+"):
            continue
        body = line[1:]
        # For Python files, strip Python comments before matching. A '#'
        # inside a string literal would still be matched — accepted false
        # positive risk, since `task.yaml` rarely lives inside strings.
        if current_file and current_file.endswith(".py"):
            comment_idx = body.find("#")
            if comment_idx >= 0:
                body = body[:comment_idx]
        for pat in _TIER_LEAK_PATTERNS:
            if pat.search(body):
                return (f"tier_leak: pattern {pat.pattern!r} introduced — "
                        f"change depends on benchmark-only metadata that won't "
                        f"exist in production. Line: {line[:120]!r}")
    return None


_EDITABLE_PATHS = (
    "agents/builder_v3.md",
    "agents/planner_v3.md",
    "agents/evaluator_v3.md",
    "harness.py",
    "bridge.py",
    "pseudo_bridge.py",
)


def _revert_working_tree() -> None:
    """Restore every tracked file in the Researcher-editable scope."""
    for rel in _EDITABLE_PATHS:
        p = ROOT / rel
        if p.exists():
            _git("checkout", "--", rel, check=False)


def _commit_iteration(label: str, loss: float) -> str:
    _git("add", *_EDITABLE_PATHS, check=False)
    _git("commit", "-m",
         f"autoresearch {label}: corpus_loss -> {loss:.4f}",
         check=False)
    return _git_sha()


# ---- Researcher agent driver ---------------------------------------------


def _build_researcher_options() -> ClaudeAgentOptions:
    """Allowed tools + system prompt for the Researcher SDK session.

    Read/Glob/Grep retained: even with editable files pre-injected per iter
    (see `_build_iter_message`), the Researcher may need to inspect files
    outside the editable scope (loss.py, store.py, eval_current.py) to
    reason about how the loop works. Those reads still incur tool-call
    overhead but they're rare relative to the editable-files reading they
    would do without pre-injection.
    """
    system_prompt = (AGENTS_DIR / "researcher.md").read_text()
    # Editable scope = system under test only:
    #   - agents/*.md (planner, builder, evaluator prompts)
    #   - harness.py (orchestration)
    #   - bridge.py / pseudo_bridge.py (builder API + impl)
    # Everything under benchmarks/ is the eval substrate and is intentionally
    # NOT in the allowed_tools — letting the Researcher edit oracle.py,
    # grader.py, loss.py, eval_current.py, or task definitions would let it
    # cheat by tuning the test rather than the system.
    allowed = [
        "Read", "Glob", "Grep",
        f"Edit({AGENTS_DIR}/builder_v3.md)",
        f"Edit({AGENTS_DIR}/planner_v3.md)",
        f"Edit({AGENTS_DIR}/evaluator_v3.md)",
        f"Edit({ROOT}/harness.py)",
        f"Edit({ROOT}/bridge.py)",
        f"Edit({ROOT}/pseudo_bridge.py)",
        f"Write({PROPOSALS_DIR}/iter_*.md)",
        # Eval invocation moved to the orchestrator — Researcher must not
        # run eval itself. These reads are still useful for context.
        "Bash(cat benchmarks/experiments/proposals/*)",
        "Bash(cat benchmarks/experiments/history.jsonl)",
        "Bash(git diff*)",
        "Bash(git status*)",
    ]
    return ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=allowed,
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
        model="sonnet",
        # Lowered from 5000 — most useful Researcher reasoning fits in ~2000 thinking tokens.
        # 5000 was leaving the Researcher to spiral into long retries on edge cases.
        max_thinking_tokens=2000,
        add_dirs=[str(ROOT / "agents"), str(ROOT / "benchmarks" / "experiments"),
                  str(ROOT / "benchmarks")],
    )


# Files whose contents we pre-inject into every per-iter message so the
# Researcher doesn't have to issue Read tool calls for them at the start
# of each turn. These are the files in editable scope — Researcher always
# reads them to understand current state before proposing edits.
_PREINJECT_PATHS = [
    "agents/builder_v3.md",
    "agents/planner_v3.md",
    "agents/evaluator_v3.md",
    "harness.py",
    "bridge.py",
    "pseudo_bridge.py",
]


def _build_iter_message(iter_n: int, baseline: dict, visible_tasks: list[str]) -> str:
    """Build the per-iter user message. Inlines the Researcher's editable
    files, prior hypotheses, and the eval command so the Researcher can
    skip the per-turn Read/inspect cycle and go straight to reasoning."""
    # Prior-iter summary so the Researcher doesn't repeat hypotheses.
    prior_summaries = []
    if HISTORY_PATH.exists():
        for line in HISTORY_PATH.read_text().splitlines():
            try:
                h = json.loads(line)
            except Exception:
                continue
            verdict = "ACCEPTED" if h.get("accepted") else "REJECTED"
            delta = h.get("paired_delta")
            delta_str = f"Δ={delta:+.4f}" if delta is not None else "Δ=n/a"
            prior_summaries.append(
                f"- iter {h.get('iter')} [{verdict}, {delta_str}]: "
                f"hypothesis={h.get('hypothesis_summary') or '(unrecorded)'!r}; "
                f"outcome={h.get('reason', '')[:120]}"
            )
    history_block = "\n".join(prior_summaries) if prior_summaries else "(no prior iters)"

    # Pre-inject current contents of every editable file. After an accept
    # these contents have changed, so re-injecting per-iter keeps the
    # Researcher's view fresh. Reading them ourselves is much cheaper
    # than ten Researcher Read tool calls.
    file_blocks = []
    for rel in _PREINJECT_PATHS:
        p = ROOT / rel
        if p.exists():
            content = p.read_text()
            file_blocks.append(f"### `{rel}`\n```\n{content}\n```")

    return (
        f"## Iteration {iter_n}\n\n"
        f"Current baseline corpus_loss: **{baseline['corpus_loss']:.4f}**\n"
        f"Per-task baseline:\n```json\n"
        f"{json.dumps(baseline.get('per_task', {}), indent=2)}\n```\n\n"
        f"### Prior iterations — hypotheses tried so far\n"
        f"{history_block}\n\n"
        f"**Do not repeat these hypotheses.** If your idea overlaps a rejected proposal, "
        f"either skip it or articulate why this new variant should differ.\n\n"
        f"### Editable files (current contents — go straight to Edit, no Read needed)\n\n"
        + "\n\n".join(file_blocks)
        + f"\n\n### Action\n\n"
        f"Proposal file: `{PROPOSALS_DIR}/iter_{iter_n}.md` (write this FIRST, before any Edit). "
        f"Lead the file with a one-sentence headline starting `# iter {iter_n}: <headline>` — "
        f"that line is recorded in history.jsonl and shown to future iters.\n"
        f"\n"
        f"**Visible tasks for this session:** {visible_tasks}\n"
        f"\n"
        f"Do one iteration now:\n"
        f"  1. Write the proposal file (hypothesis + falsification criterion).\n"
        f"  2. Apply ONE Edit to one allowed file.\n"
        f"  3. Briefly state which axis (accuracy / cost / time) and which tasks "
        f"you expect to move, and end your turn.\n"
        f"\n"
        f"**Do NOT run any eval command.** The outer driver runs the eval against "
        f"your edited working tree as soon as your turn ends, then computes the "
        f"paired-Δ promotion gate and commits or reverts. You do not see the "
        f"results until the next iter's prompt (via `history.jsonl`)."
    )


async def run_researcher_turn(client: ClaudeSDKClient, iter_n: int, baseline: dict,
                              visible_tasks: list[str]) -> str:
    """Send one iter prompt to the persistent Researcher session.

    The client lives across the whole run (created in outer_loop). Each
    iter is just a new query() on the same conversation. Anthropic's
    prompt cache stays warm on the system prompt, and the conversation's
    earlier state (which proposals were tried, how the loop responded)
    naturally carries forward without us having to re-inject.
    """
    msg = _build_iter_message(iter_n, baseline, visible_tasks)
    final_text = ""
    await client.query(msg)
    async for message in client.receive_messages():
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    final_text = block.text
        elif isinstance(message, ResultMessage):
            break
    return final_text


# ---- promotion gate -------------------------------------------------------


def _load_latest_eval_for(label: str) -> dict | None:
    conn = store_mod.connect()
    try:
        summary = store_mod.corpus_summary(conn, label)
        return summary if summary and summary.get("n") else None
    finally:
        conn.close()


def _query_pass_counts(conn, label_pattern: str, tasks: list[str],
                       latest_n: int = 3,
                       latest_n_per_tier: dict[int, int] | None = None,
                       threshold: float = PASS_ACCURACY_THRESHOLD) -> dict[str, int]:
    """Per-task count of seeds whose accuracy >= threshold.

    LIMIT per task is `latest_n_per_tier[tier]` when provided for that
    tier, else `latest_n`. Should match the tier's seed count from the
    eval that produced the rows so we count exactly that eval's cells.

    `label_pattern` accepts SQL LIKE syntax — pass `'baseline%'` for prefix
    matching across reused-baseline rows, or an exact label like
    `'iter1_visible'` for a specific iter's eval.
    """
    out: dict[str, int] = {}
    for task_id in tasks:
        n = latest_n
        if latest_n_per_tier is not None:
            n = latest_n_per_tier.get(_tier_from_task_id(task_id), latest_n)
        cur = conn.execute(
            "SELECT accuracy FROM runs WHERE label LIKE ? AND task_id = ? "
            "ORDER BY id DESC LIMIT ?",
            (label_pattern, task_id, n),
        )
        accs = [r[0] for r in cur.fetchall() if r[0] is not None]
        out[task_id] = sum(1 for a in accs if a >= threshold)
    return out


def _paired_delta(baseline: dict, candidate_label: str, tasks: list[str]) -> dict:
    """Compare candidate eval to baseline along two axes:

    1. Continuous loss delta (kept for reporting + the dashboard chart):
        delta_t       = mean(candidate_loss_t) - mean(baseline_loss_t)
        w_t           = TIER_WEIGHTS[tier_of(task_t)]   # 1.0 / 1.5 / 2.5
        corpus_delta  = Σ(w_t × delta_t) / Σ(w_t)

    2. Pass-count delta (PRIMARY GATE — promotion uses this, not 1):
        pass_t        = #seeds with accuracy >= PASS_ACCURACY_THRESHOLD
        delta_pass_t  = pass_t(candidate) - pass_t(baseline)
        pass_delta_w  = Σ(w_t × delta_pass_t)         # NOT normalized

    Why dual-track: the continuous gate is noise-dominated when accuracy
    is multi-modal across seeds (a single seed flipping pass↔fail moves
    the mean by ~0.33 on a 3-seed task). Pass-count is the discrete order
    statistic — only changes when a seed's outcome actually flips, which
    is the signal we care about.

    `baseline['label_pattern']` should be set by the caller to whatever
    sqlite label-pattern reflects the baseline rows (e.g. 'baseline%' for
    fresh/reused, the exact iter label after an iter accept rebuilds
    baseline_eval).
    """
    from benchmarks.experiments.loss import tier_weight as _tier_weight

    baseline_per_task = (baseline or {}).get("per_task") or {}
    baseline_label_pattern = (baseline or {}).get("label_pattern", "baseline%")

    conn = store_mod.connect()
    try:
        candidate_per_task = store_mod.per_task_losses(conn, candidate_label)
        baseline_pass = _query_pass_counts(
            conn, baseline_label_pattern, tasks,
            latest_n_per_tier=DEFAULT_SEEDS_PER_TIER,
        )
        candidate_pass = _query_pass_counts(
            conn, candidate_label, tasks,
            latest_n_per_tier=DEFAULT_SEEDS_PER_TIER,
        )
    finally:
        conn.close()

    rows: dict[str, dict] = {}
    weighted_deltas: list[tuple[float, float]] = []  # (loss_delta, weight)
    weighted_pass_delta = 0.0
    worst_per_task_regression = 0
    missing = []
    for task_id in tasks:
        b_row = baseline_per_task.get(task_id) or {}
        c_row = candidate_per_task.get(task_id) or {}
        b_loss = b_row.get("loss")
        c_loss = c_row.get("mean_loss")
        w = _tier_weight(task_id)
        b_pass = baseline_pass.get(task_id, 0)
        c_pass = candidate_pass.get(task_id, 0)
        delta_pass = c_pass - b_pass
        weighted_pass_delta += w * delta_pass
        if -delta_pass > worst_per_task_regression:
            worst_per_task_regression = -delta_pass

        if b_loss is None or c_loss is None:
            missing.append(task_id)
            rows[task_id] = {"baseline_loss": b_loss, "new_loss": c_loss,
                             "delta": None, "weight": w,
                             "baseline_pass": b_pass, "new_pass": c_pass,
                             "delta_pass": delta_pass}
            continue
        d = c_loss - b_loss
        weighted_deltas.append((d, w))
        rows[task_id] = {"baseline_loss": round(b_loss, 4),
                         "new_loss": round(c_loss, 4),
                         "delta": round(d, 4),
                         "weight": w,
                         "baseline_pass": b_pass,
                         "new_pass": c_pass,
                         "delta_pass": delta_pass,
                         "n_candidate": c_row.get("n")}

    if weighted_deltas:
        wsum = sum(w for _, w in weighted_deltas)
        corpus_delta = sum(d * w for d, w in weighted_deltas) / wsum
    else:
        corpus_delta = None

    return {
        "corpus_delta": corpus_delta,
        "pass_count_delta_weighted": round(weighted_pass_delta, 2),
        "worst_per_task_pass_regression": worst_per_task_regression,
        "per_task": rows,
        "missing_in_candidate": missing,
        "n_tasks_compared": len(weighted_deltas),
    }


async def _run_holdout(holdout_tasks: list[str]) -> dict[str, Any]:
    log.info(f"running holdout gate on {holdout_tasks}…")
    return await run_eval(
        tasks=holdout_tasks,
        seeds=3,
        seeds_per_tier=DEFAULT_SEEDS_PER_TIER,
        label=f"holdout_{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}",
        model="sonnet",
        skip_planner=False,
        time_budget=HOLDOUT_TIME_BUDGET,
        max_turns=60,
        parallel=DEFAULT_PARALLEL,
    )


async def _run_iter_visible(iter_n: int, visible_tasks: list[str]) -> dict[str, Any]:
    """Run the per-iter visible eval. Owned by the orchestrator (not the
    Researcher) so that (a) Researcher session can end as soon as it's
    done editing, eliminating the stale-background-task coordination bug,
    and (b) we can apply the tighter VISIBLE_TIME_BUDGET cap consistently."""
    log.info(f"running iter {iter_n} visible eval on {visible_tasks}…")
    return await run_eval(
        tasks=visible_tasks,
        seeds=3,
        seeds_per_tier=DEFAULT_SEEDS_PER_TIER,
        label=f"iter{iter_n}_visible",
        model="sonnet",
        skip_planner=False,
        time_budget=VISIBLE_TIME_BUDGET,
        max_turns=60,
        parallel=DEFAULT_PARALLEL,
    )


def _append_history(entry: dict) -> None:
    HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
    with HISTORY_PATH.open("a") as f:
        f.write(json.dumps(entry, default=str) + "\n")


def _try_reuse_baseline(label_prefix: str, tasks: list[str],
                        min_seeds: int = 2,
                        min_seeds_per_tier: dict[int, int] | None = None) -> dict | None:
    """Reconstruct a baseline eval report from existing store rows.

    Returns None if we don't have enough coverage to avoid a fresh run.
    Per-task minimum is `min_seeds_per_tier[tier]` when provided for that
    tier, otherwise `min_seeds`. Reads the raw result.json off disk to
    recompute corpus_loss via the current loss function — so any recent
    loss tweak is reflected in the reused number.
    """
    from benchmarks.experiments.loss import corpus_loss as _corpus_loss

    def _required(task_id: str) -> int:
        if min_seeds_per_tier is None:
            return min_seeds
        return min_seeds_per_tier.get(_tier_from_task_id(task_id), min_seeds)

    conn = store_mod.connect()
    try:
        rows: list[dict] = []
        for task_id in tasks:
            need = _required(task_id)
            cur = conn.execute(
                """SELECT result_json_path FROM runs
                    WHERE label LIKE ? AND task_id = ?
                    ORDER BY id DESC LIMIT ?""",
                (f"{label_prefix}%", task_id, need),
            )
            cells = [dict(r) for r in cur.fetchall()]
            if len(cells) < need:
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


# ---- process-group management --------------------------------------------


def _write_pidfile() -> None:
    """Own a fresh process group and record our PIDs for shutdown.py.

    Claude Agent SDK subprocesses + their eval_current children + Excel
    instances all inherit this PGID unless they explicitly setsid, so
    `kill -TERM -<pgid>` cleans up the whole tree in one shot. Without
    this, SIGKILL to the main process orphans subprocesses and leaves
    Excel + HTTP servers running.
    """
    try:
        os.setpgrp()
    except OSError as e:
        log.warning(f"could not create new process group: {e}")
    payload = {
        "pid": os.getpid(),
        "pgid": os.getpgrp(),
        "started_at": datetime.now(timezone.utc).isoformat(),
    }
    PIDFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    PIDFILE_PATH.write_text(json.dumps(payload, indent=2))
    log.info(f"pid={payload['pid']} pgid={payload['pgid']} pidfile={PIDFILE_PATH}")


def _remove_pidfile() -> None:
    try:
        PIDFILE_PATH.unlink(missing_ok=True)
    except Exception:
        pass


def _install_signal_handlers() -> None:
    """Catch TERM/INT to mark shutdown and propagate to the process group.

    We kill the group ourselves rather than relying on the shell — ensures
    Excel/xlwings children die even if they've been re-parented.
    """
    def _handler(signum, frame):
        log.warning(f"received signal {signum}; tearing down…")
        _write_status(phase="shutting_down", received_signal=signum)
        _remove_pidfile()
        try:
            os.killpg(os.getpgrp(), signal.SIGTERM)
        except Exception:
            pass
        # Escalate if anything survives after a short grace window.
        time.sleep(2)
        try:
            os.killpg(os.getpgrp(), signal.SIGKILL)
        except Exception:
            pass

    # Ignore SIGHUP. This loop runs for hours unattended; if the launching
    # terminal closes (or the IDE/Claude session that spawned it exits),
    # Python's default SIGHUP action would kill the process without
    # running our handler — leaving status frozen and the pidfile
    # orphaned. Explicit shutdown happens via shutdown.py (SIGTERM).
    try:
        signal.signal(signal.SIGHUP, signal.SIG_IGN)
    except (ValueError, OSError):
        pass

    for sig in (signal.SIGTERM, signal.SIGINT):
        try:
            signal.signal(sig, _handler)
        except (ValueError, OSError):
            pass


# ---- outer loop -----------------------------------------------------------


async def outer_loop(*, max_iters: int, max_dollars: float,
                     skip_holdout: bool = False,
                     rotate: bool = False,
                     seed: int | None = None) -> None:
    PROPOSALS_DIR.mkdir(parents=True, exist_ok=True)
    _write_pidfile()
    _install_signal_handlers()

    if not _git_working_tree_clean():
        _remove_pidfile()
        raise SystemExit(
            "refuse to start autoresearch with a dirty working tree. "
            "commit or stash first."
        )

    # Wipe stale `iter*_visible` and `iter*_canary` rows from sqlite. Each
    # session's iter labels (iter1_visible, iter2_visible, …) collide with
    # any prior session's iter labels — `_load_latest_eval_for` aggregates
    # ALL rows under a label, so without this wipe the new iter 1's
    # paired-Δ would be polluted by the prior session's iter 1 rows.
    # Baseline / holdout rows are preserved (they're keyed by label
    # prefix, not iter number, and reusable across sessions).
    conn = store_mod.connect()
    try:
        cur = conn.execute(
            "DELETE FROM runs WHERE label LIKE 'iter%_visible' "
            "OR label LIKE 'iter%_canary'"
        )
        conn.commit()
        if cur.rowcount:
            log.info(f"purged {cur.rowcount} stale iter*_visible/canary rows from prior session")
    finally:
        conn.close()

    # Pick the task sets for this session. Rotation is fixed-per-session:
    # the Researcher sees the same visible set across every iter so its
    # proposals are comparable; rotation only happens between sessions.
    if rotate:
        visible_tasks, holdout_tasks = rotate_sets(seed)
        log.info(f"rotate=True seed={seed} — visible={visible_tasks} holdout={holdout_tasks}")
    else:
        visible_tasks, holdout_tasks = list(VISIBLE_SET), list(HOLDOUT_SET)
        log.info(f"rotate=False — using fixed sets. visible={visible_tasks}")

    _write_status(phase="baseline", iter=0, started_at=datetime.now(timezone.utc).isoformat(),
                  max_iters=max_iters, max_dollars=max_dollars,
                  visible_tasks=visible_tasks, holdout_tasks=holdout_tasks,
                  seed=seed)

    # Baseline: try to reuse existing store rows before spending on a fresh eval.
    # Reuse threshold tracks DEFAULT_SEEDS_PER_TIER so each tier has the same
    # sample count it would get in a fresh run (paired-Δ noise scales with
    # mismatched seed counts between baseline and candidate).
    baseline_eval = _try_reuse_baseline(
        "baseline", visible_tasks,
        min_seeds_per_tier=DEFAULT_SEEDS_PER_TIER,
    )
    if baseline_eval is not None:
        log.info(f"reusing stored baseline — corpus_loss={baseline_eval['corpus_loss']:.4f} "
                 f"(n={baseline_eval['n_runs']} over {len(visible_tasks)} tasks)")
    else:
        log.info(f"no usable stored baseline — running fresh on visible set {visible_tasks}…")
        baseline_eval = await run_eval(
            tasks=visible_tasks, seeds=3, seeds_per_tier=DEFAULT_SEEDS_PER_TIER,
            label="baseline",
            model="sonnet", skip_planner=False,
            time_budget=VISIBLE_TIME_BUDGET, max_turns=60, parallel=DEFAULT_PARALLEL,
        )
    # Track the sqlite label-pattern for pass-count queries in _paired_delta.
    # Both the fresh-run and reuse paths produce/draw rows under "baseline%"
    # (the reuse query uses LIKE prefix, fresh-run writes label="baseline").
    baseline_eval["label_pattern"] = "baseline%"
    LAST_EVAL_PATH.write_text(json.dumps(baseline_eval, indent=2, default=str))
    baseline_loss = baseline_eval["corpus_loss"]
    log.info(f"baseline corpus_loss = {baseline_loss:.4f}")

    # Holdout baseline — only needed if we'll gate. Try reuse first.
    holdout_baseline = None
    if not skip_holdout:
        _write_status(phase="holdout_baseline")
        holdout_baseline = _try_reuse_baseline(
            "holdout_", holdout_tasks,
            min_seeds_per_tier=DEFAULT_SEEDS_PER_TIER,
        )
        if holdout_baseline is not None:
            log.info(f"reusing stored holdout baseline — "
                     f"corpus_loss={holdout_baseline['corpus_loss']:.4f}")
        else:
            log.info(f"no usable stored holdout — running fresh on {holdout_tasks}…")
            holdout_baseline = await _run_holdout(holdout_tasks)
        log.info(f"baseline holdout_loss = {holdout_baseline['corpus_loss']:.4f}")

    spent = (baseline_eval.get("eval_cost_usd") or 0.0) + \
            ((holdout_baseline.get("eval_cost_usd") if holdout_baseline else 0.0) or 0.0)
    current_sha = _git_sha()
    accepted_since_last_holdout = 0

    # Persistent Researcher SDK session — single connect for the whole run.
    # Per-iter messages reuse the same conversation so Anthropic's prompt
    # cache stays warm on the system prompt + system tools list. The
    # conversation history naturally carries forward (no need to re-inject
    # what was tried). If conversation outgrows the context window over
    # many iters, we'll see it in errors and add a periodic reset.
    researcher_client = ClaudeSDKClient(options=_build_researcher_options())
    await researcher_client.connect()
    log.info("researcher session connected (persistent across iters)")

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
            verdict_text = await run_researcher_turn(researcher_client, i, baseline_eval, visible_tasks)
        except Exception as e:
            log.warning(f"researcher turn crashed: {type(e).__name__}: {e}")

        # Orchestrator owns the eval (Researcher used to invoke it via Bash,
        # which produced "stale background task" coordination bugs and made
        # the per-iter wall = think + eval ≈ 13 min). Now: Researcher writes
        # proposal + applies edit + exits; orchestrator runs eval against
        # the edited tree.
        _write_status(phase="evaluating", iter=i)
        visible_label = f"iter{i}_visible"
        try:
            await _run_iter_visible(i, visible_tasks)
        except Exception as e:
            log.warning(f"iter {i} visible eval crashed: {type(e).__name__}: {e}")
        visible_summary = _load_latest_eval_for(visible_label)

        accepted = False
        reason = "no_eval"
        new_loss = None
        ran_holdout_this_iter = False

        # Tier-leak lint: reject before paired-delta math if the diff
        # introduces branches keyed on benchmark-only metadata. Saves
        # eval cost on iters that wouldn't be safe to commit anyway.
        leak_reason = _check_for_tier_leaks(branch)
        paired = None
        failure_diag = None
        if leak_reason is not None:
            reason = leak_reason
            failure_diag = leak_reason
            log.warning(f"iter {i}: {leak_reason}")
        else:
            # Paired-by-task comparison: per-task delta, averaged. Much less
            # noisy than raw corpus_loss comparison.
            paired = _paired_delta(baseline_eval, visible_label, visible_tasks) if visible_summary else None

            # Diagnose any failure mode separately so history shows the real cause.
            if not visible_summary:
                failure_diag = "no_eval_data_in_store (researcher's eval subprocess didn't insert rows — usually a hang or crash)"
            elif paired is None or paired.get("corpus_delta") is None:
                missing = (paired or {}).get("missing_in_candidate") or []
                failure_diag = f"paired_delta_uncomputable (missing per-task data; missing_tasks={missing})"

        if leak_reason is None and visible_summary and paired and paired["corpus_delta"] is not None:
            new_loss = visible_summary["loss"]
            corpus_delta = paired["corpus_delta"]
            pass_delta = paired["pass_count_delta_weighted"]
            worst_reg = paired["worst_per_task_pass_regression"]
            gate_summary = (f"pass-Δ={pass_delta:+.1f}, worst-task-regression={worst_reg}, "
                            f"loss-Δ={corpus_delta:+.4f}")

            # Two-path accept criterion. Both paths share the per-task
            # regression guard (no individual task lost ≥2 seeds).
            #   primary  = pass-rate gate     → real failure-mode resolutions
            #   secondary = continuous big-win → cost/speed wins, or
            #               sub-threshold accuracy gains too large to be noise
            primary_pass = pass_delta >= MIN_PASS_COUNT_DELTA_WEIGHTED
            secondary_pass = (
                pass_delta >= 0
                and corpus_delta <= -CONTINUOUS_BIG_WIN_THRESHOLD
            )

            if paired["missing_in_candidate"]:
                reason = (f"incomplete eval — missing tasks in candidate: "
                          f"{paired['missing_in_candidate']}")
            elif worst_reg > MAX_PER_TASK_REGRESSION:
                reason = (f"per_task_regression (one task lost {worst_reg} seeds, "
                          f"max allowed {MAX_PER_TASK_REGRESSION}; {gate_summary})")
            elif not (primary_pass or secondary_pass):
                reason = (
                    f"no_improvement (need pass-Δ ≥ {MIN_PASS_COUNT_DELTA_WEIGHTED} "
                    f"OR (pass-Δ ≥ 0 AND loss-Δ ≤ -{CONTINUOUS_BIG_WIN_THRESHOLD}); "
                    f"{gate_summary})"
                )
            else:
                gate_path = "pass-rate" if primary_pass else "continuous"
                if skip_holdout:
                    accepted = True
                    reason = f"improved via {gate_path} ({gate_summary}); holdout skipped"
                else:
                    would_be_accepted_count = accepted_since_last_holdout + 1
                    if would_be_accepted_count >= HOLDOUT_EVERY_N_ACCEPTED:
                        _write_status(phase="holdout_gate", iter=i)
                        holdout_new = await _run_holdout(holdout_tasks)
                        ran_holdout_this_iter = True
                        spent += holdout_new.get("eval_cost_usd", 0.0) or 0.0
                        if (holdout_new["corpus_loss"]
                                > holdout_baseline["corpus_loss"] + HOLDOUT_REGRESSION_MAX):
                            reason = (f"holdout regression after {would_be_accepted_count} accepted — "
                                      f"{holdout_new['corpus_loss']:.4f} > "
                                      f"{holdout_baseline['corpus_loss']:.4f} + {HOLDOUT_REGRESSION_MAX}")
                        else:
                            accepted = True
                            reason = (f"improved via {gate_path} ({gate_summary}) + holdout ok "
                                      f"(after {would_be_accepted_count} accepted)")
                            holdout_baseline = holdout_new
                            accepted_since_last_holdout = 0
                    else:
                        accepted = True
                        reason = (f"improved via {gate_path} ({gate_summary}); holdout due in "
                                  f"{HOLDOUT_EVERY_N_ACCEPTED - would_be_accepted_count} more accepts")
                        accepted_since_last_holdout = would_be_accepted_count
        else:
            reason = failure_diag or "researcher did not produce a visible eval"

        # Extract the Researcher's hypothesis headline BEFORE the reject path's
        # `git clean -fd` wipes the proposal file. Future iters read this from
        # history.jsonl to avoid re-proposing the same idea.
        hypothesis_summary = None
        proposal_path = PROPOSALS_DIR / f"iter_{i}.md"
        if proposal_path.exists():
            text = proposal_path.read_text().strip()
            for raw in text.splitlines():
                ln = raw.strip()
                if ln:
                    hypothesis_summary = ln.lstrip("#").strip()[:240]
                    break
        if not hypothesis_summary and verdict_text:
            for raw in verdict_text.splitlines():
                ln = raw.strip()
                if ln:
                    hypothesis_summary = ln.lstrip("#").strip()[:240]
                    break

        if accepted:
            sha = _commit_iteration(f"iter{i}", new_loss)
            baseline_loss = new_loss
            # Rebuild per_task from the candidate's stored rows. Earlier we
            # used `dict(visible_summary)`, which is the FLAT corpus_summary
            # ({n, loss, accuracy, ...}) — not a {task_id: {...}} map.
            # That broke _paired_delta on every subsequent iter because
            # baseline_per_task.get(task_id) returned None for every task.
            conn = store_mod.connect()
            try:
                accepted_per_task = store_mod.per_task_losses(conn, visible_label)
            finally:
                conn.close()
            baseline_eval = {
                "corpus_loss": new_loss,
                "per_task": {
                    t: {
                        "n": row.get("n"),
                        "loss": row.get("mean_loss"),
                        "mean_accuracy": row.get("mean_accuracy"),
                        "mean_cost_cold_usd": row.get("mean_cost_cold_usd"),
                        "mean_wall_seconds": row.get("mean_wall_seconds"),
                    }
                    for t, row in accepted_per_task.items()
                },
                # Pass-count queries in the next iter's _paired_delta need
                # to read from the exact iter-visible label (not the original
                # 'baseline%' prefix), since this iter's accepted rows ARE
                # the new baseline.
                "label_pattern": visible_label,
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

        _append_history({
            "iter": i,
            "ts": datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "accepted": accepted,
            "reason": reason,
            "hypothesis_summary": hypothesis_summary,
            "baseline_loss": baseline_loss,
            "new_loss": new_loss,
            "paired_delta": (paired or {}).get("corpus_delta") if paired else None,
            "pass_count_delta_weighted": (paired or {}).get("pass_count_delta_weighted") if paired else None,
            "worst_per_task_pass_regression": (paired or {}).get("worst_per_task_pass_regression") if paired else None,
            "paired_per_task": (paired or {}).get("per_task") if paired else None,
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

    # Tear down the persistent Researcher session.
    try:
        await researcher_client.disconnect()
    except Exception as e:
        log.warning(f"researcher disconnect: {type(e).__name__}: {e}")

    _write_status(phase="done", baseline_loss=baseline_loss, spent_usd=round(spent, 4))
    log.info(f"done. spent=${spent:.2f}, baseline_loss={baseline_loss:.4f}")
    _remove_pidfile()


# ---- CLI -----------------------------------------------------------------


def _main() -> int:
    fmt = logging.Formatter(
        "%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    # File handler — always on so a record survives terminal close /
    # parent death even if the console handler's pipe is broken. The
    # last run that died on SIGHUP left no log on disk, which made
    # post-mortem painful.
    log_path = BENCH_ROOT / "experiments" / "autoresearch.log"
    log_path.parent.mkdir(parents=True, exist_ok=True)
    fh = logging.FileHandler(log_path, mode="a")
    fh.setFormatter(fmt)
    root.addHandler(fh)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    root.addHandler(sh)
    log.info(f"logging to file: {log_path}")
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.autoresearch")
    ap.add_argument("--max-iters", type=int, default=60)
    ap.add_argument("--max-dollars", type=float, default=400.0)
    ap.add_argument("--skip-holdout", action="store_true",
                    help="Skip the holdout gate. Only for debugging the loop itself.")
    ap.add_argument("--rotate", action="store_true",
                    help="Rotate in fresh visible + holdout subsets from the pool "
                         "(eval_current.VISIBLE_POOL / HOLDOUT_POOL). Deterministic given --seed. "
                         "Sets stay fixed within a session so iter-to-iter comparisons remain valid.")
    ap.add_argument("--seed", type=int, default=None,
                    help="Seed for --rotate. Default: wall-clock time so each run sees new tasks.")
    args = ap.parse_args()
    seed = args.seed if args.seed is not None else int(time.time())
    asyncio.run(outer_loop(
        max_iters=args.max_iters,
        max_dollars=args.max_dollars,
        skip_holdout=args.skip_holdout,
        rotate=args.rotate,
        seed=seed,
    ))
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
