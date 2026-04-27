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
    HOLDOUT_SET, VISIBLE_SET, run_eval, resolve_task_set, rotate_sets,
)


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
IMPROVEMENT_EPSILON   = 0.015  # paired delta must be at least this negative.
# Calibrated for the rewritten log-scale loss (loss.py). Per-cell loss is
# now bounded ~[0, 2.5] (was ~[0, 2.6]), but typical per-cell values are
# 0.05–0.5 instead of 0.4–1.0 — old 0.03 threshold was 6% of typical loss
# range; 0.015 keeps the same 6% in the new scale.
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
DEFAULT_PARALLEL = 4


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
    of task_meta inside files we own (we read tier in headless_builder.py
    just for the result.json metadata) don't false-positive — those lines
    aren't ADDED in this iter's diff.

    Comment-aware: for .py files, content past a `#` is stripped before
    matching, so a Researcher comment like
    "# Production-safe: not task.yaml metadata" doesn't false-positive.
    """
    scope = [
        "agents/builder_v3.md",
        "agents/planner_v3.md",
        "benchmarks/headless_builder.py",
        "benchmarks/pseudo_bridge.py",
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
    "benchmarks/headless_builder.py",
    "benchmarks/pseudo_bridge.py",
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

    tasks_csv = ",".join(visible_tasks)
    eval_cmd = (
        f"python -m benchmarks.experiments.eval_current "
        f"--tasks {tasks_csv} --seeds 3 --label iter{iter_n}_visible"
    )

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
        f"**Eval command to run (exactly this):**\n"
        f"```\n{eval_cmd}\n```\n"
        f"\n"
        f"Do one iteration now: write the proposal, apply ONE edit, run the eval "
        f"command above, and finish with your ACCEPT/REJECT recommendation."
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


def _paired_delta(baseline: dict, candidate_label: str, tasks: list[str]) -> dict:
    """Compare candidate eval to baseline using per-task paired deltas.

    For each task in `tasks`:
        delta_t = mean(candidate_loss_t) - mean(baseline_loss_t)
    Then corpus_delta = mean of delta_t across tasks.

    Task-level means cancel out the (huge) variance from task difficulty,
    so this number is MUCH more stable than (candidate_flat_mean - baseline_flat_mean).

    Returns {
        "corpus_delta": float,
        "per_task": {task: {baseline_loss, new_loss, delta}},
        "missing_in_candidate": [task_id,...],  # tasks with no candidate rows
    }
    """
    baseline_per_task = (baseline or {}).get("per_task") or {}
    conn = store_mod.connect()
    try:
        candidate_per_task = store_mod.per_task_losses(conn, candidate_label)
    finally:
        conn.close()

    rows: dict[str, dict] = {}
    deltas: list[float] = []
    missing = []
    for task_id in tasks:
        b_row = baseline_per_task.get(task_id) or {}
        c_row = candidate_per_task.get(task_id) or {}
        b_loss = b_row.get("loss")
        c_loss = c_row.get("mean_loss")
        if b_loss is None or c_loss is None:
            missing.append(task_id)
            rows[task_id] = {"baseline_loss": b_loss, "new_loss": c_loss, "delta": None}
            continue
        d = c_loss - b_loss
        deltas.append(d)
        rows[task_id] = {"baseline_loss": round(b_loss, 4),
                         "new_loss": round(c_loss, 4),
                         "delta": round(d, 4),
                         "n_candidate": c_row.get("n")}

    return {
        "corpus_delta": (sum(deltas) / len(deltas)) if deltas else None,
        "per_task": rows,
        "missing_in_candidate": missing,
        "n_tasks_compared": len(deltas),
    }


async def _run_holdout(holdout_tasks: list[str]) -> dict[str, Any]:
    log.info(f"running holdout gate on {holdout_tasks}…")
    return await run_eval(
        tasks=holdout_tasks,
        seeds=3,
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
    # min_seeds=1 is intentional — trading a bit of noise for ~15min of re-eval time.
    baseline_eval = _try_reuse_baseline("baseline", visible_tasks, min_seeds=1)
    if baseline_eval is not None:
        log.info(f"reusing stored baseline — corpus_loss={baseline_eval['corpus_loss']:.4f} "
                 f"(n={baseline_eval['n_runs']} over {len(visible_tasks)} tasks)")
    else:
        log.info(f"no usable stored baseline — running fresh on visible set {visible_tasks}…")
        baseline_eval = await run_eval(
            tasks=visible_tasks, seeds=3, label="baseline",  # 3 seeds for variance control (see paired comparison)
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
        holdout_baseline = _try_reuse_baseline("holdout_", holdout_tasks, min_seeds=1)
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

        _write_status(phase="evaluating", iter=i)

        # Fresh eval is the Researcher's last `iter{i}_visible` batch.
        # Canary is no longer run — visible is small enough to be the primary gate.
        visible_label = f"iter{i}_visible"
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
            # Require full task coverage — if a task is missing a candidate run,
            # we don't have enough data to trust the paired delta.
            if paired["missing_in_candidate"]:
                reason = (f"incomplete eval — missing tasks in candidate: "
                          f"{paired['missing_in_candidate']}")
            elif corpus_delta > -IMPROVEMENT_EPSILON:
                reason = (f"no_improvement (paired Δ={corpus_delta:+.4f}, "
                          f"need ≤ -{IMPROVEMENT_EPSILON})")
            elif skip_holdout:
                accepted = True
                reason = f"improved paired Δ={corpus_delta:+.4f}, holdout skipped"
            else:
                # Tentatively accept on visible; only fire holdout every Nth accept.
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
                        reason = (f"improved (paired Δ={corpus_delta:+.4f}) + holdout ok "
                                  f"(after {would_be_accepted_count} accepted)")
                        holdout_baseline = holdout_new
                        accepted_since_last_holdout = 0
                else:
                    accepted = True
                    reason = (f"improved (paired Δ={corpus_delta:+.4f}; holdout due in "
                              f"{HOLDOUT_EVERY_N_ACCEPTED - would_be_accepted_count} more accepts)")
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
    ap.add_argument("--max-iters", type=int, default=10)
    ap.add_argument("--max-dollars", type=float, default=50.0)
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
