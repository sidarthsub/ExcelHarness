"""Headless builder driver.

Ties together:
  - pseudo_bridge (xlwings-backed local HTTP bridge)
  - Claude Agent SDK Builder (sonnet, writes scripts against bridge.py)
  - Oracle (auto-answers Planner clarifications; currently skipped — bench
    tasks come with a full brief + context + seed, the Planner phase is
    handed by providing the spec up front)
  - grader (scores the saved xlsx)

Flow:
  1. Allocate a run directory under benchmarks/runs/<timestamp>/<task_id>/
  2. Start PseudoBridgeServer on an ephemeral port pointed at
     <run_dir>/candidate/model.xlsx (optionally copying a task stub)
  3. Spawn the Builder with BRIDGE_URL set; feed it the task brief + inputs
     + seed clarifications as if the Planner had already resolved them.
  4. Builder writes python scripts in <run_dir>/scripts/ that import
     bridge.Bridge and drive the pseudo-bridge. It calls b.checkpoint()
     between sheets (no eval feedback in headless mode — just saves).
  5. When Builder emits "Model complete. Ready for review." OR the time
     budget is exceeded, the orchestrator shuts down the server.
  6. Run the grader on the final xlsx and emit a JSON result.

The Builder never knows it's not live Excel — same prompt, same bridge.py.
"""
from __future__ import annotations

import asyncio
import json
import logging
import os
import shutil
import socket
import threading
import time
import yaml

import psutil
from datetime import datetime
from pathlib import Path
from typing import Any

import re

from claude_agent_sdk import ClaudeSDKClient, ClaudeAgentOptions
from claude_agent_sdk.types import (
    AssistantMessage,
    ResultMessage,
    TextBlock,
    ToolUseBlock,
)

import jsonschema

from benchmarks.grader import grade
from benchmarks.oracle import OracleConfig, format_clarifications, resolve_questions
from benchmarks.pseudo_bridge import PseudoBridgeServer

log = logging.getLogger("headless_builder")

ROOT = Path(__file__).resolve().parent.parent
BENCH_ROOT = Path(__file__).resolve().parent
AGENTS_DIR = ROOT / "agents"
TASKS_ROOT = BENCH_ROOT / "tasks"
RUNS_ROOT = BENCH_ROOT / "runs"


# ---- helpers ---------------------------------------------------------------


class _CellWatchdog:
    """Hard wall-time enforcement at the OS level.

    asyncio.wait_for is supposed to enforce the time_budget at the loop level,
    but during network outages the Claude Agent SDK's receive_messages() can
    park on a subprocess stdin/stdout read that doesn't surface
    CancelledError cleanly — observed on 2026-04-23 when wifi loss left a
    cell stuck for 2h+ past its 900s budget. This watchdog runs in a
    daemon thread independent of the event loop and SIGKILLs the SDK's
    subprocess tree when the deadline passes. Closing the SDK's stdout
    is what finally unblocks the asyncio coroutine.

    Targets only PIDs that appeared between snapshot() and arm() — i.e.,
    the subprocess(es) this cell spawned. Other concurrent cells are
    untouched.
    """

    def __init__(self, label: str, deadline_seconds: float):
        self.label = label
        self.deadline_seconds = deadline_seconds
        self._before: set[int] = set()
        self._target_pids: set[int] = set()
        self._timer: threading.Timer | None = None
        self._fired = False

    @staticmethod
    def _children() -> set[int]:
        try:
            return {p.pid for p in psutil.Process().children(recursive=True)}
        except Exception:
            return set()

    def snapshot(self) -> None:
        """Call BEFORE spawning the SDK client subprocess."""
        self._before = self._children()

    def arm(self) -> None:
        """Call AFTER the SDK client has spawned. Captures new PIDs and starts the timer."""
        after = self._children()
        self._target_pids = after - self._before
        log.info(f"watchdog[{self.label}]: tracking pids={sorted(self._target_pids)} "
                 f"deadline={self.deadline_seconds:.0f}s")
        self._timer = threading.Timer(self.deadline_seconds, self._fire)
        self._timer.daemon = True
        self._timer.start()

    def _fire(self) -> None:
        self._fired = True
        log.warning(f"watchdog[{self.label}] FIRED — SIGKILL on {sorted(self._target_pids)}")
        for pid in self._target_pids:
            try:
                proc = psutil.Process(pid)
                # kill the whole subtree of this PID too
                for child in proc.children(recursive=True):
                    try: child.kill()
                    except psutil.NoSuchProcess: pass
                proc.kill()
            except psutil.NoSuchProcess:
                pass
            except Exception as e:
                log.warning(f"watchdog[{self.label}] kill {pid} failed: {e}")

    def cancel(self) -> bool:
        """Disarm the timer. Returns True if the watchdog had already fired."""
        if self._timer is not None:
            self._timer.cancel()
        return self._fired


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _read_text(p: Path) -> str:
    return p.read_text() if p.exists() else ""


def _build_planner_context_text(task_dir: Path) -> str:
    """All inline context the Planner gets: brief + every input document."""
    parts = [f"## Brief\n{_read_text(task_dir / 'brief.md')}"]
    inputs_dir = task_dir / "inputs"
    if inputs_dir.exists():
        for f in sorted(inputs_dir.iterdir()):
            if f.is_file() and f.suffix.lower() in (".md", ".txt"):
                parts.append(f"## Input document: {f.name}\n{_read_text(f)}")
    return "\n\n---\n\n".join(parts)


def _extract_json_array(text: str) -> list | None:
    """Pull the first JSON array out of a Planner Pass-1 response."""
    text = text.strip()
    # Fenced code block
    m = re.search(r"```(?:json)?\s*(\[[\s\S]*?\])\s*```", text)
    if m:
        try:
            return json.loads(m.group(1))
        except json.JSONDecodeError:
            pass
    # Bare array
    m = re.search(r"\[[\s\S]*\]", text)
    if m:
        try:
            return json.loads(m.group(0))
        except json.JSONDecodeError:
            pass
    return None


async def run_planner_phase(
    task_id: str,
    task_dir: Path,
    run_dir: Path,
    *,
    model: str = "sonnet",
) -> tuple[dict | None, dict]:
    """Run Planner (Pass 1 + Oracle + Pass 2) and return (spec, stats).

    spec is the validated model_spec.json contents, or None if spec generation
    failed (we fall back to passing the brief straight to the Builder).
    stats contains planner_cost, oracle_stats, question counts.
    """
    log.info("running Planner phase (with Oracle)…")
    planner_prompt = (AGENTS_DIR / "planner_v3.md").read_text()
    schema_path = ROOT / "schemas" / "model_spec.schema.json"
    schema = json.loads(schema_path.read_text())

    spec_path = run_dir / "model_spec.json"
    opts = ClaudeAgentOptions(
        system_prompt=planner_prompt,
        allowed_tools=[
            "Read", "Glob", "Grep",
            f"Write({spec_path})", f"Edit({spec_path})",
        ],
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
        model=model,
        add_dirs=[str(task_dir), str(run_dir)],
    )
    client = ClaudeSDKClient(options=opts)

    stats = {
        "planner_cost_usd": 0.0,
        "planner_turns": 0,
        "planner_input_tokens": 0,
        "planner_output_tokens": 0,
        "oracle": None,
        "questions_asked": 0,
        "clarifications": {},
    }

    async def _query(msg: str, label: str) -> str:
        nonlocal stats
        await client.query(msg)
        final_text = ""
        async for message in client.receive_messages():
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        final_text = block.text
            elif isinstance(message, ResultMessage):
                u = message.usage or {}
                stats["planner_input_tokens"] += u.get("input_tokens", 0) or 0
                stats["planner_output_tokens"] += u.get("output_tokens", 0) or 0
                stats["planner_cost_usd"] += message.total_cost_usd or 0.0
                stats["planner_turns"] += 1
                break
        log.info(f"planner.{label}: {len(final_text)} chars")
        return final_text

    ctx_text = _build_planner_context_text(task_dir)

    try:
        await client.connect()

        # Pass 1: ambiguity detection
        pass1_msg = (
            f"SCHEMA: {schema_path}\n\n"
            f"--- CONTEXT ---\n{ctx_text}\n---\n\n"
            "Run Pass 1 (ambiguity detection). Output a JSON array of questions, "
            "or an empty array if the brief is fully specified. Respond with ONLY the JSON array."
        )
        pass1_text = await _query(pass1_msg, "pass1")
        questions = _extract_json_array(pass1_text) or []
        stats["questions_asked"] = len(questions)
        log.info(f"planner asked {len(questions)} question(s)")

        # Oracle resolves
        clarifications: dict[str, str] = {}
        if questions:
            seed_path = task_dir / "clarifications.seed.yaml"
            task_meta = yaml.safe_load((task_dir / "task.yaml").read_text())
            budget = (task_meta.get("clarification_budget") or {})
            cfg = OracleConfig(
                task_id=task_id,
                context_md=_read_text(task_dir / "context.md"),
                seed_path=seed_path if seed_path.exists() else None,
                max_questions=budget.get("max_questions", 10),
            )
            answers, oracle_stats = await resolve_questions(questions, cfg)
            clarifications = answers
            stats["oracle"] = oracle_stats
            log.info(
                f"oracle: {oracle_stats['seed_hits']} seed, "
                f"{oracle_stats['cache_hits']} cached, "
                f"{oracle_stats['llm_calls']} llm calls"
            )
        stats["clarifications"] = clarifications

        # Pass 2: spec
        clarif_block = "\n".join(f"- {k}: {v}" for k, v in clarifications.items()) or "(none)"
        pass2_msg = (
            f"CLARIFICATIONS:\n{clarif_block}\n\n"
            f"Run Pass 2. Write the full spec to {spec_path} using the Write tool, then output the JSON."
        )
        pass2_text = await _query(pass2_msg, "pass2")

        # Load + validate
        if not spec_path.exists():
            # Try extracting from response
            m = re.search(r"```json\s*(\{[\s\S]*?\})\s*```", pass2_text)
            if not m:
                m = re.search(r"(\{[\s\S]*\"sheets\"[\s\S]*\})", pass2_text)
            if m:
                extracted = m.group(1) if m.lastindex else m.group(0)
                spec_path.write_text(extracted)
        if not spec_path.exists():
            log.warning("planner did not produce a spec — falling back to brief-only builder")
            return None, stats

        try:
            spec = json.loads(spec_path.read_text())
            jsonschema.validate(instance=spec, schema=schema)
        except (json.JSONDecodeError, jsonschema.ValidationError) as e:
            log.warning(f"planner spec invalid ({e}); attempting one fix pass")
            await _query(
                f"Your spec failed validation: {e}. Fix it and save again to {spec_path}.",
                "pass2_fix",
            )
            try:
                spec = json.loads(spec_path.read_text())
                jsonschema.validate(instance=spec, schema=schema)
            except Exception as e2:
                log.warning(f"spec still invalid after fix ({e2}); using brief-only builder")
                return None, stats

        return spec, stats
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def _results_contract_block(task_meta: dict, candidate_dir: Path) -> str:
    """Build the results.json contract section for the Builder prompt.

    If the task defines output_keys in task.yaml, the Builder MUST produce
    a results.json file next to model.xlsx mapping each key to the
    Sheet!Cell address that contains that value. Type hints:
      - "number"  → the resolved cell's value will be compared numerically
      - "formula" → the value is compared AND the cell must contain a formula
      - "text"    → reserved for future text checks
    """
    keys = task_meta.get("output_keys") or []
    if not keys:
        return ""
    results_path = candidate_dir / "results.json"
    lines = [
        "## Required output artifact: results.json",
        "",
        f"Before emitting the completion sentinel, write the file `{results_path}`",
        "as a JSON object mapping each semantic key to the `Sheet!Cell` address",
        "containing that value in the workbook. Example:",
        "",
        "```json",
        "{",
        '  "pre_money_valuation": "Inputs!B4",',
        '  "post_money_valuation": "Inputs!B6"',
        "}",
        "```",
        "",
        "Required keys:",
        "",
    ]
    for k in keys:
        kind = k.get("type", "number")
        desc = k.get("description", "")
        note = {"number": "numeric value compared with tolerance",
                "formula": "numeric value compared AND cell must contain a formula",
                "text": "text match (reserved)"}.get(kind, "")
        lines.append(f"- **{k['key']}** ({kind}) — {desc}  [_{note}_]")
    lines.append("")
    lines.append("Do NOT write this file until all sheets are built and saved. The Builder")
    lines.append("bridge's `b.read_values()` is a reliable way to confirm the cell you're pointing at.")
    lines.append("")
    return "\n\n---\n\n".join(["\n".join(lines), ""])


def _describe_stub(stub_path: Path) -> str:
    """Describe the stub's contents as text so the Builder can recreate it.

    The pseudo-bridge always starts from a blank workbook (opening stubs
    via xlwings hangs on macOS). We compensate by telling the Builder what
    the stub looked like: sheet names, which cells are already filled, and
    their values. The Builder writes those first, then proceeds to the task.
    """
    import openpyxl
    try:
        wb = openpyxl.load_workbook(stub_path, data_only=False)
    except Exception as e:
        return f"(could not load stub {stub_path.name}: {e})"
    lines = [f"## Stub starting state: {stub_path.name}",
             "The workbook starts BLANK — reproduce these pre-filled cells as your first step:"]
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        lines.append(f"\n### Sheet: `{sheet_name}`")
        if ws.max_row == 1 and ws.max_column == 1 and ws["A1"].value in (None, ""):
            lines.append("  (empty)")
            continue
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None or v == "":
                    continue
                fmt = cell.number_format
                fmt_note = f"  [{fmt}]" if fmt and fmt != "General" else ""
                lines.append(f"  {cell.coordinate}: {v!r}{fmt_note}")
    wb.close()
    return "\n".join(lines)


def build_builder_context(task_dir: Path, task_meta: dict, spec: dict | None = None) -> str:
    """Assemble a single prompt block the Builder gets up front.

    Contains:
      - the validated spec from the Planner phase (if available), OR
        the brief directly (fallback when Planner fails)
      - a text description of the stub (pseudo-bridge starts blank)
      - all reference .md/.txt input documents
      - absolute paths to .xlsx input files
    """
    parts: list[str] = []
    if spec is not None:
        parts.append(f"## model_spec.json (from Planner)\n{json.dumps(spec, indent=2)}")
    else:
        parts.append(f"## Task brief\n{_read_text(task_dir / 'brief.md')}")
    stub = task_meta.get("stub_file")
    if stub:
        stub_path = task_dir / stub
        if stub_path.exists():
            parts.append(_describe_stub(stub_path))
    inputs_dir = task_dir / "inputs"
    if inputs_dir.exists():
        for f in sorted(inputs_dir.iterdir()):
            if f.is_file() and f.suffix.lower() in (".md", ".txt"):
                parts.append(f"## Input document: {f.name}\n{_read_text(f)}")
        xlsx_list = [str(f) for f in sorted(inputs_dir.glob("*.xlsx"))]
        if xlsx_list:
            parts.append(
                "## Input xlsx files (use absolute paths with bridge.copy_sheet_from_input)\n"
                + "\n".join(xlsx_list)
            )
    return "\n\n---\n\n".join(parts)


# ---- the driver ------------------------------------------------------------


async def run_headless(
    task_id: str,
    *,
    model: str = "sonnet",
    time_budget_seconds: float | None = None,
    max_turns: int = 40,
    skip_planner: bool = False,
) -> dict:
    task_dir = TASKS_ROOT / task_id
    if not task_dir.exists():
        raise FileNotFoundError(f"task {task_id} not found in {TASKS_ROOT}")
    task_meta = yaml.safe_load((task_dir / "task.yaml").read_text())
    time_budget = time_budget_seconds or task_meta.get("time_budget_seconds", 600)

    # Run dir — microsecond precision + short random suffix so concurrent
    # cells for the SAME task never collide on stamp-only paths. Previously
    # two cells firing in the same second would share a run_dir, and the
    # second one's writes would clobber the first — losing a cell's results
    # at store-insertion time.
    import uuid
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    suffix = uuid.uuid4().hex[:6]
    run_dir = RUNS_ROOT / f"{stamp}_{task_id}_{suffix}"
    scripts_dir = run_dir / "scripts"
    candidate_dir = run_dir / "candidate"
    for d in (run_dir, scripts_dir, candidate_dir):
        d.mkdir(parents=True, exist_ok=True)
    candidate_xlsx = candidate_dir / "model.xlsx"

    port = _pick_free_port()
    t_start = time.time()

    usage_totals = {"input_tokens": 0, "output_tokens": 0,
                    "cache_read_input_tokens": 0, "cache_creation_input_tokens": 0,
                    "cost_usd": 0.0, "turns": 0}

    completed = False
    terminated_reason = "unknown"

    # --- Planner + Oracle phase ---
    spec: dict | None = None
    planner_stats: dict = {}
    task_tier = task_meta.get("tier", 0)
    # Tier-0 tasks are atomic and fully specified via results_contract +
    # brief; skip the planner to avoid 2 extra sonnet turns (~$0.10 cold).
    _should_skip_planner = skip_planner or (task_tier == 0)
    if not _should_skip_planner:
        try:
            spec, planner_stats = await run_planner_phase(task_id, task_dir, run_dir, model=model)
        except Exception as e:
            log.warning(f"planner phase failed: {type(e).__name__}: {e}")
            planner_stats = {"error": f"{type(e).__name__}: {e}"}

    with PseudoBridgeServer(
        output_xlsx=candidate_xlsx, port=port, host="127.0.0.1"
    ) as server:
        # BRIDGE_URL is passed per-run via ClaudeAgentOptions.env below.
        # Do NOT write to os.environ — process-global state would race
        # between concurrent run_headless() calls and every Builder would
        # end up pointed at whichever bridge wrote last.

        system_prompt = (AGENTS_DIR / "builder_v3.md").read_text()

        # Tier-0 and tier-1 tasks are sufficiently simple for haiku, which is
        # 3× cheaper per input token than sonnet. Tier-1 tasks like
        # inputs_from_term_sheet are mechanical extraction tasks (read doc →
        # copy values) that do not require sonnet-level reasoning. Extended
        # thinking is a sonnet-only feature, so disable it when downgrading.
        builder_model = model
        builder_thinking_tokens = 5000
        if task_tier <= 1 and model == "sonnet":
            builder_model = "haiku"
            builder_thinking_tokens = 0

        opts = ClaudeAgentOptions(
            system_prompt=system_prompt,
            allowed_tools=[
                "Read", "Glob", "Grep",
                f"Write({scripts_dir}/builder_*.py)",
                f"Edit({scripts_dir}/builder_*.py)",
                f"Write({candidate_dir}/results.json)",
                f"Edit({candidate_dir}/results.json)",
                f"Bash(python3 {scripts_dir}/builder_*.py)",
                "Bash(ls*)", "Bash(cat*)",
            ],
            permission_mode="bypassPermissions",
            cwd=str(ROOT),
            model=builder_model,
            max_thinking_tokens=builder_thinking_tokens,
            add_dirs=[str(task_dir), str(run_dir)],
            env={
                "BRIDGE_URL": server.base_url,
                "BRIDGE_VERIFY_TLS": "0",
            },
        )

        client = ClaudeSDKClient(options=opts)
        context = build_builder_context(task_dir, task_meta, spec=spec)
        candidate_abs = str(candidate_xlsx)

        # If the task specifies output_keys, inject the results.json contract so
        # the Builder emits a machine-readable map of semantic keys → cell addresses.
        results_contract = _results_contract_block(task_meta, candidate_dir)
        initial_msg = (
            f"You are building an Excel model in a **headless** pseudo-bridge context.\n"
            f"- The bridge is at {server.base_url} (plain HTTP). Default Bridge() picks it up via env.\n"
            f"- Construct: `b = Bridge()` — no args needed.\n"
            f"- Write one script per sheet to {scripts_dir}/builder_<SheetName>.py, then run it.\n"
            f"- The workbook auto-saves to {candidate_abs} on every b.checkpoint() and at server shutdown.\n"
            f"- When the model is done, emit the sentinel exactly:\n"
            f"  Model complete. Ready for review.\n"
            f"- NO multi-sheet eval feedback loop here. b.checkpoint() just saves.\n"
            f"  If you want a sanity check, Read the saved xlsx or call b.dump_sheet().\n"
            f"- HEADLESS SINGLE-PASS: Do NOT emit any between-sheet yield sentences (e.g. 'Sheet 1 done, moving to...'). "
            f"Build all sheets back-to-back in the fewest turns possible. There are no fail files or mid-session eval "
            f"messages to wait for. Proceed directly from one sheet to the next and emit the final sentinel only "
            f"when ALL sheets are complete and results.json is written.\n"
            f"- The task ships with pre-resolved clarifications below.\n\n"
            f"{results_contract}"
            f"{context}"
        )

        # Hard wall-time enforcement at the OS level. asyncio.wait_for can
        # fail to cancel through the SDK's subprocess reads when the network
        # is flaky; this watchdog SIGKILLs the SDK subprocess tree when the
        # deadline passes — closing stdout, which unblocks the async read.
        watchdog = _CellWatchdog(label=task_id, deadline_seconds=time_budget + 60)

        async def run_turns():
            nonlocal completed, terminated_reason, usage_totals
            watchdog.snapshot()
            await client.connect()
            watchdog.arm()
            try:
                await client.query(initial_msg)
                for turn in range(1, max_turns + 1):
                    if time.time() - t_start > time_budget:
                        terminated_reason = "time_budget_exceeded"
                        return
                    final_text = ""
                    tool_uses = []
                    async for message in client.receive_messages():
                        if isinstance(message, AssistantMessage):
                            for block in message.content:
                                if isinstance(block, TextBlock):
                                    final_text = block.text
                                elif isinstance(block, ToolUseBlock):
                                    tool_uses.append(block.name)
                        elif isinstance(message, ResultMessage):
                            u = message.usage or {}
                            usage_totals["input_tokens"] += u.get("input_tokens", 0) or 0
                            usage_totals["output_tokens"] += u.get("output_tokens", 0) or 0
                            usage_totals["cache_read_input_tokens"] += u.get("cache_read_input_tokens", 0) or 0
                            usage_totals["cache_creation_input_tokens"] += u.get("cache_creation_input_tokens", 0) or 0
                            usage_totals["cost_usd"] += message.total_cost_usd or 0.0
                            usage_totals["turns"] += 1
                            break

                    log.info(
                        f"turn {turn}: {len(tool_uses)} tools ({', '.join(tool_uses[:4])}{'…' if len(tool_uses)>4 else ''}) "
                        f"— text={len(final_text)} chars"
                    )

                    if "Model complete. Ready for review." in final_text:
                        completed = True
                        terminated_reason = "sentinel"
                        return

                    # Time check before asking for another turn
                    if time.time() - t_start > time_budget:
                        terminated_reason = "time_budget_exceeded"
                        return

                    # Prompt for next turn — mirrors the live harness "continue" fallthrough
                    await client.query("continue")
                terminated_reason = "max_turns"
            finally:
                try:
                    await client.disconnect()
                except Exception:
                    pass

        try:
            # Soft timeout: asyncio.wait_for cancels run_turns at +30s past budget.
            await asyncio.wait_for(run_turns(), timeout=time_budget + 30)
        except asyncio.TimeoutError:
            terminated_reason = "time_budget_exceeded"
        finally:
            # Hard timeout fallback: if the watchdog fired before we reached
            # this finally, the SDK subprocess was force-killed at the OS
            # level. Surface that distinct reason so it shows up in history.
            if watchdog.cancel():
                terminated_reason = "watchdog_killed_wall_budget"

    wall = time.time() - t_start

    # Grade
    graded: dict = {"accuracy": 0.0, "passed": 0, "total": 0,
                    "weighted_score": 0.0, "total_weight": 0.0, "checks": []}
    grading_err: str | None = None
    if candidate_xlsx.exists():
        try:
            graded = grade(
                candidate_xlsx=candidate_xlsx,
                grading_yaml=task_dir / "gold" / "grading.yaml",
                task_dir=task_dir,
                recalc=True,
            )
        except Exception as e:
            grading_err = f"{type(e).__name__}: {e}"
    else:
        grading_err = "candidate xlsx not produced"

    total_cost = usage_totals["cost_usd"] + planner_stats.get("planner_cost_usd", 0.0)
    if planner_stats.get("oracle"):
        total_cost += planner_stats["oracle"].get("cost_dollars", 0.0)
    result = {
        "task_id": task_id,
        "tier": task_meta.get("tier"),
        "cost_budget_dollars": task_meta.get("cost_budget_dollars"),
        "run_dir": str(run_dir),
        "candidate": str(candidate_xlsx),
        "spec_generated": spec is not None,
        "completed": completed,
        "terminated_reason": terminated_reason,
        "wall_seconds": round(wall, 2),
        "time_budget_seconds": time_budget,
        "over_budget": wall > time_budget,
        "builder_model": builder_model,
        "builder_usage": usage_totals,
        "planner_stats": planner_stats,
        "dollars": round(total_cost, 4),
        "accuracy": graded.get("accuracy"),
        "passed": graded.get("passed"),
        "total": graded.get("total"),
        "weighted_score": graded.get("weighted_score"),
        "total_weight": graded.get("total_weight"),
        "grading_error": grading_err,
        "checks": graded.get("checks"),
    }
    (run_dir / "result.json").write_text(json.dumps(result, indent=2, default=str))
    return result


# ---- CLI ------------------------------------------------------------------


def _main() -> None:
    import argparse
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", required=True)
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--time-budget", type=float, default=None,
                    help="Override task.yaml time budget (seconds).")
    ap.add_argument("--max-turns", type=int, default=40)
    ap.add_argument("--pretty", action="store_true")
    ap.add_argument("--skip-planner", action="store_true",
                    help="Skip the Planner+Oracle phase; feed the brief to the Builder directly.")
    args = ap.parse_args()

    result = asyncio.run(run_headless(
        task_id=args.task,
        model=args.model,
        time_budget_seconds=args.time_budget,
        max_turns=args.max_turns,
        skip_planner=args.skip_planner,
    ))
    # Strip per-check detail from the stdout summary unless --pretty
    if args.pretty:
        print(json.dumps(result, indent=2, default=str))
    else:
        summary = {k: v for k, v in result.items() if k != "checks"}
        print(json.dumps(summary, default=str))


if __name__ == "__main__":
    _main()
