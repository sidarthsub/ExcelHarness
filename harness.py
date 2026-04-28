#!/usr/bin/env python3
"""ExcelHarness v3 — single-process live-Excel harness.

Flow:
  1. Start PseudoBridgeServer (HTTPS + WSS).
  2. Wait for add-in connection.
  3. Read brief from user via chat.
  4. Run Planner phase (ambiguity Q&A -> spec -> self-review).
  5. (future) Run Builder phase.
  6. (future) Run Evaluator at checkpoints.
"""
from __future__ import annotations

import asyncio
import json
import logging
import re
import sys
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)-7s %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("harness")

from pseudo_bridge import PseudoBridgeServer
from benchmarks.oracle import OracleConfig, resolve_questions
from session import Session

from claude_agent_sdk import query, ClaudeSDKClient
from claude_agent_sdk.types import (
    AssistantMessage,
    ClaudeAgentOptions,
    ResultMessage,
    TextBlock,
    ToolUseBlock,
)
import jsonschema

ROOT = Path(__file__).parent
AGENTS_DIR = ROOT / "agents"

# Session-level usage accumulator. Populated by _track_usage() on every
# ResultMessage from any agent (planner, builder, evaluator); summed at exit.
_SESSION_USAGE: dict[str, dict] = {}


def _track_usage(agent: str, label: str, result: ResultMessage) -> None:
    """Log per-call usage and accumulate into the session total."""
    usage = result.usage or {}
    cost = result.total_cost_usd or 0.0
    input_t = usage.get("input_tokens", 0)
    output_t = usage.get("output_tokens", 0)
    cache_r = usage.get("cache_read_input_tokens", 0)
    cache_w = usage.get("cache_creation_input_tokens", 0)

    log.info(
        f"usage[{agent}/{label}]: in={input_t:,} out={output_t:,} "
        f"cache_r={cache_r:,} cache_w={cache_w:,} ${cost:.4f}"
    )

    bucket = _SESSION_USAGE.setdefault(
        agent, {"input": 0, "output": 0, "cache_r": 0, "cache_w": 0, "cost_usd": 0.0, "calls": 0}
    )
    bucket["input"] += input_t
    bucket["output"] += output_t
    bucket["cache_r"] += cache_r
    bucket["cache_w"] += cache_w
    bucket["cost_usd"] += cost
    bucket["calls"] += 1


def _log_session_totals() -> None:
    """Dump the accumulated usage by agent plus the grand total."""
    if not _SESSION_USAGE:
        return
    grand = 0.0
    log.info("=" * 60)
    log.info("Session usage totals:")
    for agent, b in sorted(_SESSION_USAGE.items()):
        log.info(
            f"  {agent:10s} calls={b['calls']:<3} in={b['input']:>10,} "
            f"out={b['output']:>9,} cache_r={b['cache_r']:>11,} "
            f"cache_w={b['cache_w']:>10,} ${b['cost_usd']:.4f}"
        )
        grand += b["cost_usd"]
    log.info(f"  {'TOTAL':10s} ${grand:.4f}")
    log.info("=" * 60)


async def run_agent_session(
    system_prompt: str,
    user_messages: list[str],
    allowed_tools: list[str],
    cwd: Path,
    on_activity: callable | None = None,
    add_dirs: list[str] | None = None,
    model: str | None = None,
    usage_label: str = "call",
    usage_agent: str = "agent",
) -> str:
    """Run a single-turn Claude Agent SDK query and return the text of the final assistant response.

    on_activity: optional async callback(str) called with status updates as the
    agent works (tool calls, progress). Used to stream updates to chat so the
    user isn't staring at a blank screen.

    usage_label / usage_agent are emitted into the per-call usage log line
    and the session totals.
    """
    options = ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=allowed_tools,
        permission_mode="bypassPermissions",
        cwd=str(cwd),
        add_dirs=add_dirs or [],
        model=model,
    )

    combined = "\n\n---\n\n".join(user_messages)

    final_text = ""
    async for message in query(prompt=combined, options=options):
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    final_text = block.text
                elif isinstance(block, ToolUseBlock) and on_activity:
                    # Only surface meaningful tool calls — skip noisy reads/greps
                    tool_name = block.name
                    tool_input = block.input or {}
                    if tool_name == "Write":
                        path = tool_input.get("file_path", "")
                        short = path.split("/")[-1] if "/" in str(path) else path
                        await on_activity(f"Writing {short}...")
                    elif tool_name == "Bash":
                        cmd = str(tool_input.get("command", ""))
                        # Only show python script executions, not trivial commands
                        if "python3" in cmd and "builder" in cmd:
                            await on_activity(f"Running build script...")
        elif isinstance(message, ResultMessage):
            _track_usage(usage_agent, usage_label, message)
    return final_text


def extract_pdf_inputs(session: Session) -> list[Path]:
    """Extract text from every *.pdf in input/ and write a companion *.txt.

    Skips PDFs whose companion .txt already exists and is newer than the PDF.
    The downstream planner context already globs *.txt, so the extracted
    files are picked up automatically.
    """
    import pypdf
    input_dir = session.input_dir
    written: list[Path] = []
    for pdf in sorted(input_dir.glob("*.pdf")):
        txt = pdf.with_suffix(".txt")
        if txt.exists() and txt.stat().st_mtime >= pdf.stat().st_mtime:
            continue
        try:
            reader = pypdf.PdfReader(str(pdf))
            chunks = []
            for i, page in enumerate(reader.pages, 1):
                chunks.append(f"--- Page {i} ---\n{(page.extract_text() or '').strip()}")
            txt.write_text("\n\n".join(chunks))
            written.append(txt)
            log.info(f"Extracted {pdf.name} -> {txt.name} ({len(reader.pages)} pages)")
        except Exception as e:
            log.warning(f"PDF extract failed on {pdf.name}: {e}")
    return written


def preprocess_inputs(session: Session, sheet_filter: dict | None = None) -> list[Path]:
    """Run dump.py on each xlsx. If sheet_filter provided, only dump selected sheets per file."""
    import subprocess, shutil
    input_dir = session.input_dir
    xlsx_files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xls"))
    if not xlsx_files:
        return []

    dumps_dir = input_dir / "dumps"
    dumps_dir.mkdir(exist_ok=True)
    evals_dir = ROOT / "evals"

    for xlsx in xlsx_files:
        sheets = sheet_filter.get(xlsx) if sheet_filter else None
        desc = f" ({len(sheets)} sheets)" if sheets else " (all)"
        log.info(f"Dumping {xlsx.name}{desc}...")
        if evals_dir.exists():
            shutil.rmtree(evals_dir)
        cmd = [sys.executable, str(ROOT / "dump.py"), str(xlsx)]
        if sheets:
            cmd += ["--sheets", ",".join(sheets)]
        result = subprocess.run(
            cmd, capture_output=True, text=True, cwd=str(ROOT),
        )
        if result.returncode != 0:
            log.warning(f"dump.py failed on {xlsx.name}: {result.stderr[:200]}")
            continue

        file_dumps = dumps_dir / xlsx.stem
        file_dumps.mkdir(exist_ok=True)
        for subdir in ["formulas", "styles", "screenshots"]:
            src = evals_dir / subdir
            if src.exists():
                dst = file_dumps / subdir
                if dst.exists():
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)

    if evals_dir.exists():
        shutil.rmtree(evals_dir)
    return xlsx_files


def _dedupe_styles(dump_dirs: list[Path]) -> list[Path]:
    """From all style dumps, pick one representative per sheet type.
    Groups by category (Series A, Series B, Returns, etc.) and picks the best.
    """
    import re
    groups: dict[str, list[Path]] = {}
    for dump_dir in dump_dirs:
        styles_dir = dump_dir / "styles"
        if not styles_dir.exists():
            continue
        for f in sorted(styles_dir.glob("*.txt")):
            name = f.stem
            # Remove source tag: "[VulnCheck...] Series A (1011 Leads)" -> "Series A (1011 Leads)"
            name = re.sub(r'^\[.*?\]\s*', '', name)
            # Categorize: "Series A (...)" -> "Series A", "Post A Returns ..." -> "Post A Returns"
            # Extract the core type
            if name.startswith("Series A"):
                cat = "Series A"
            elif name.startswith("Series B"):
                cat = "Series B"
            elif name.startswith("Series C"):
                cat = "Series C"
            elif name.startswith("Series Seed"):
                cat = "Series Seed"
            elif "Post A Returns" in name:
                cat = "Post A Returns"
            elif "Post B Returns" in name:
                cat = "Post B Returns"
            elif "Post C Returns" in name:
                cat = "Post C Returns"
            elif "Returns Analysis" in name:
                cat = "Returns Analysis"
            elif "Cap Table" in name or "Cap Post" in name:
                cat = f"Cap Table: {name}"  # keep cap tables separate
            else:
                cat = name  # unique — keep as-is
            groups.setdefault(cat, []).append(f)

    result = []
    for cat, files in sorted(groups.items()):
        if len(files) == 1:
            result.append(files[0])
        else:
            # Prefer "1011 Leads" variant, else largest
            picked = next((f for f in files if "1011 Leads" in f.name), None)
            if not picked:
                picked = max(files, key=lambda f: f.stat().st_size)
            result.append(picked)
    return result


def build_context_for_planner(session: Session, brief: str) -> str:
    """Consolidated context for Planner — injected into prompt, no Read calls needed."""
    parts = [f"## Brief\n{brief}"]

    # Include any text/pdf input files (SAFE docs, term sheets, etc.)
    for f in sorted(session.input_dir.glob("*.txt")):
        parts.append(f"## Input document: {f.name}\n{f.read_text()}")

    # Include formula dumps (may be absent if task has no xlsx inputs)
    dumps_root = session.input_dir / "dumps"
    if dumps_root.exists():
        for dump_dir in sorted(dumps_root.iterdir()):
            formulas_dir = dump_dir / "formulas"
            if not formulas_dir.exists():
                continue
            for f in sorted(formulas_dir.glob("*.txt")):
                parts.append(f"## Formula dump: {f.name}\n{f.read_text()}")
    return "\n\n---\n\n".join(parts)


def build_context_for_builder(session: Session, spec: dict) -> str:
    """Consolidated context for Builder — injected into prompt, no Read calls needed."""
    parts = []
    parts.append(f"## model_spec.json\n{json.dumps(spec, indent=2)}")

    # Q&A clarifications are summarized into spec constraints by the Planner

    parts.append(f"## bridge.py\n{(ROOT / 'bridge.py').read_text()}")

    # Reference dumps — may be absent if the task has no xlsx inputs.
    dumps_root = session.input_dir / "dumps"
    dump_dirs: list[Path] = sorted(dumps_root.iterdir()) if dumps_root.exists() else []

    for f in _dedupe_styles(dump_dirs):
        parts.append(f"## Style: {f.name}\n{f.read_text()}")

    for dump_dir in dump_dirs:
        formulas_dir = dump_dir / "formulas"
        if formulas_dir.exists():
            for f in sorted(formulas_dir.glob("*.txt")):
                parts.append(f"## Formulas: {f.name}\n{f.read_text()}")

    screenshots = []
    for dump_dir in dump_dirs:
        ss_dir = dump_dir / "screenshots"
        if ss_dir.exists():
            for f in sorted(ss_dir.glob("*.jpg")) + sorted(ss_dir.glob("*.png")):
                screenshots.append(str(f))
    if screenshots:
        parts.append("## Reference Screenshots (use Read to view)\n" + "\n".join(screenshots))

    return "\n\n---\n\n".join(parts)


async def run_planner(session: Session, oracle_cfg: OracleConfig, brief: str,
                      planner_context: str | None = None) -> tuple[dict, dict]:
    """Planner using ClaudeSDKClient — single session, all passes share context.

    Returns (spec, oracle_stats). oracle_stats is the resolve_questions
    output: {seed_hits, cache_hits, llm_calls, over_budget, tokens_in,
    tokens_out, cost_dollars}.
    """
    prompt = (AGENTS_DIR / "planner_v3.md").read_text()
    schema = json.loads((ROOT / "schemas" / "model_spec.schema.json").read_text())
    spec_path = session.run_dir / "model_spec.json"

    options = ClaudeAgentOptions(
        system_prompt=prompt,
        allowed_tools=["Read", "Glob", "Grep", f"Write({session.run_dir}/model_spec.json)", f"Edit({session.run_dir}/model_spec.json)"],
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
    )
    client = ClaudeSDKClient(options=options)

    async def get_response(label: str, msg: str) -> str:
        """Send a message and collect the text response."""
        mark(f"{label}_send")
        await client.query(msg)
        final_text = ""
        tool_count = 0
        async for message in client.receive_messages():
            if isinstance(message, AssistantMessage):
                for block in message.content:
                    if isinstance(block, TextBlock):
                        final_text = block.text
                    elif isinstance(block, ToolUseBlock):
                        tool_count += 1
                        log.info(f"Planner tool: {block.name}")
            elif isinstance(message, ResultMessage):
                _track_usage("planner", label, message)
                break
        mark(f"{label}_done ({tool_count} tools, {len(final_text)} chars)")
        return final_text

    try:
        from timing import mark
        await client.connect()
        mark("planner_connected")

        # --- Pass 1: Ambiguities ---
        log.info("planner: analyzing brief for ambiguities")

        pass1_msg = f"SCHEMA: {ROOT / 'schemas' / 'model_spec.schema.json'}\n\n"
        if planner_context:
            pass1_msg += f"--- CONTEXT (brief + all input data) ---\n{planner_context}\n---\n\n"
        pass1_msg += (
            f"BRIEF:\n{brief}\n\n"
            "Run Pass 1 (ambiguity detection). Output a JSON array of questions, "
            "or an empty array if the brief is fully specified. Respond with ONLY the JSON array."
        )

        pass1_text = await get_response("pass1", pass1_msg)
        mark("pass1_complete")

        try:
            questions = json.loads(pass1_text.strip())
        except json.JSONDecodeError:
            m = re.search(r"\[.*\]", pass1_text, re.DOTALL)
            questions = json.loads(m.group(0)) if m else []

        # --- Resolve all clarifications via the oracle in a single batch ---
        # In production this is a chat-driven Q&A; in headless the oracle reads
        # the task's clarifications.seed.yaml + context.md and answers all
        # questions at once. The interactive per-answer pruning loop is moot
        # here because oracle resolution is atomic.
        clarifications: dict[str, str] = {}
        oracle_stats: dict = {"seed_hits": 0, "cache_hits": 0, "llm_calls": 0,
                              "over_budget": 0, "tokens_in": 0, "tokens_out": 0,
                              "cost_dollars": 0.0}
        if questions:
            log.info(f"planner: {len(questions)} clarifications → oracle")
            answers, oracle_stats = await resolve_questions(questions, oracle_cfg)
            for qid, ans in answers.items():
                clarifications[qid] = ans
                session.append_clarification(qid, ans)
            log.info(f"oracle: seed={oracle_stats['seed_hits']} cache={oracle_stats['cache_hits']} "
                     f"llm={oracle_stats['llm_calls']} over_budget={oracle_stats['over_budget']}")

        # --- Pass 2: Full spec (same session — context cached) ---
        mark(f"qa_complete ({len(clarifications)} answers)")
        log.info("planner: generating spec")
        clarif_text = "\n".join(f"- {k}: {v}" for k, v in clarifications.items())
        pass2_msg = (
            f"CLARIFICATIONS:\n{clarif_text or '(none)'}\n\n"
            "Run Pass 2. Write the full spec and save it to "
            f"{spec_path} using the Write tool. Then output the JSON."
        )

        pass2_text = await get_response("pass2", pass2_msg)
        mark("pass2_complete")
        log.info(f"Planner Pass 2 response: {len(pass2_text)} chars, spec exists: {spec_path.exists()}")

        # --- Load, validate ---
        if not spec_path.exists():
            # Try to extract JSON from response text
            m = re.search(r"```json\s*(\{.*?\})\s*```", pass2_text, re.DOTALL)
            if not m:
                m = re.search(r"(\{[\s\S]*\"sheets\"[\s\S]*\})", pass2_text, re.DOTALL)
            if m:
                extracted = m.group(1) if m.lastindex else m.group(0)
                spec_path.write_text(extracted)
                log.info("planner: spec extracted from response (Write tool didn't fire)")
            else:
                log.info(f"Pass 2 text preview: {pass2_text[:500]}")
                raise RuntimeError("Planner did not write model_spec.json and no JSON found in response")

        spec = json.loads(spec_path.read_text())
        try:
            jsonschema.validate(instance=spec, schema=schema)
        except jsonschema.ValidationError as e:
            log.info(f"planner: spec failed validation: {e.message}, fixing")
            fix_text = await get_response("pass2_fix", f"Your spec failed validation: {e.message}\n\nFix it and save again.")
            spec = json.loads(spec_path.read_text())
            jsonschema.validate(instance=spec, schema=schema)

        # --- Pass 3: Self-review (same session — has full context) ---
        mark("pass2_validated")
        log.info("planner: self-review")
        pass3_text = await get_response(
            "pass3",
            "Run Pass 3 (self-review). Check for contradictions, overprescription, "
            "incompleteness, or redundancy. If you find issues, rewrite the spec. "
            "If it's clean, reply 'CLEAN'."
        )

        # Reload in case pass 3 rewrote
        spec = json.loads(spec_path.read_text())
        jsonschema.validate(instance=spec, schema=schema)

        return spec, oracle_stats

    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


async def prepare_eval_input(
    session: Session,
    server: PseudoBridgeServer,
    checkpoint_description: str,
    spec: dict,
    turn: int,
) -> tuple[Path, bytes]:
    """Snapshot the workbook, dump each sheet, render screenshots, and write them into an eval_input dir.

    Returns a tuple of (eval_dir, xlsx_bytes). The bytes are returned so the
    caller can also use them for git commits without re-reading the file.
    """
    import base64
    import shutil
    from snapshot_renderer import render_xlsx_to_pngs
    from live_dump import dump_xlsx_to_json

    eval_dir = session.run_dir / "eval_input"
    # Clean and recreate each checkpoint so the evaluator sees only current state.
    if eval_dir.exists():
        shutil.rmtree(eval_dir)
    eval_dir.mkdir(parents=True)
    (eval_dir / "dumps").mkdir()
    (eval_dir / "screenshots").mkdir()

    # Checkpoint description + spec
    (eval_dir / "checkpoint.md").write_text(checkpoint_description)
    (eval_dir / "spec.json").write_text(json.dumps(spec, indent=2))

    # In headless, pseudo_bridge auto-saves the candidate xlsx on every
    # b.checkpoint() call (see PseudoBridgeServer worker loop). Just read
    # the live file rather than RPCing for a snapshot.
    candidate_xlsx = Path(server.output_xlsx)
    xlsx_bytes = b""
    snap_path: Path | None = None
    if candidate_xlsx.exists():
        xlsx_bytes = candidate_xlsx.read_bytes()
        snap_path = session.snapshots_dir / f"turn_{turn}.xlsx"
        snap_path.write_bytes(xlsx_bytes)
    else:
        log.warning(f"candidate xlsx not yet on disk: {candidate_xlsx}")

    # Render PNGs (may be skipped if snap missing or renderer fails — evaluator
    # still has dumps + spec but its prompt requires screenshots).
    if snap_path is not None:
        try:
            render_xlsx_to_pngs(snap_path, eval_dir / "screenshots")
        except Exception as e:
            log.warning(f"Snapshot render failed: {e}")

    # Dump each sheet from the snapshot xlsx via openpyxl.
    if snap_path is not None:
        try:
            written = dump_xlsx_to_json(snap_path, eval_dir / "dumps")
            log.info(f"Dumped {len(written)} sheet(s) from snapshot.")
        except Exception as e:
            log.error(f"Snapshot dump failed: {e}")
    else:
        log.warning("No snapshot bytes — skipping sheet dump.")

    # Reference styles are read directly from session.input_dir at eval
    # time — no need to re-copy per checkpoint.

    return eval_dir, xlsx_bytes


async def run_evaluator(eval_dir: Path, input_dumps_dir: Path) -> dict:
    """Fresh query() per checkpoint — clean context, thorough review.

    input_dumps_dir is session.input_dir / "dumps" — the harness reads
    reference style files directly from the original input dumps instead of
    copying them into eval_dir on every checkpoint.
    """
    prompt = (AGENTS_DIR / "evaluator_v3.md").read_text()

    # Build eval context inline — minimize Read calls
    parts = [f"## Checkpoint\n{(eval_dir / 'checkpoint.md').read_text()}"]
    parts.append(f"## Spec\n{(eval_dir / 'spec.json').read_text()}")

    dumps_dir = eval_dir / "dumps"
    if dumps_dir.exists():
        for f in sorted(dumps_dir.glob("*.json")):
            parts.append(f"## Dump: {f.stem}\n{f.read_text()}")

    if input_dumps_dir.exists():
        for dump_dir in sorted(input_dumps_dir.iterdir()):
            styles_dir = dump_dir / "styles"
            if styles_dir.exists():
                for f in sorted(styles_dir.glob("*.txt")):
                    parts.append(f"## Reference Style: {f.stem}\n{f.read_text()}")

    # List screenshots for Read (images can't be injected as text).
    # All pages are listed — cap removed so the evaluator can always find
    # the sheet being checkpointed (previously [:2] showed only the first
    # two pages, which was often the wrong sheet).
    ss_dir = eval_dir / "screenshots"
    ss_files = []
    if ss_dir.exists():
        ss_files = [f for f in sorted(ss_dir.iterdir()) if f.suffix in (".jpg", ".jpeg", ".png")]
        if ss_files:
            parts.append(
                "## SCREENSHOTS — Read these to check visual formatting\n"
                + "\n".join(str(f) for f in ss_files)
            )

    context = "\n\n---\n\n".join(parts)
    user_msg = (
        f"Review this checkpoint.\n"
        f"ALL text context is below. You MUST also Read the {len(ss_files)} screenshot image(s) listed at the bottom to verify formatting.\n"
        f"Return ONLY a JSON object per your output format spec.\n\n"
        f"{context}"
    )

    async def eval_activity(msg: str):
        log.info(f"Evaluator: {msg}")

    result_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[user_msg],
        allowed_tools=["Read", "Glob", "Grep"],
        cwd=ROOT,
        on_activity=eval_activity,
        model="sonnet",
        usage_agent="evaluator",
        usage_label=eval_dir.parent.name,
    )

    m = re.search(r"```json\s*(\{.*?\})\s*```", result_text, re.DOTALL)
    if not m:
        m = re.search(r"\{.*\}", result_text, re.DOTALL)
    if not m:
        return {
            "status": "fail",
            "findings": [{"severity": "error", "issue": "Evaluator did not return JSON"}],
        }
    try:
        verdict = json.loads(m.group(1) if m.lastindex else m.group(0))
    except json.JSONDecodeError as e:
        return {
            "status": "fail",
            "findings": [{"severity": "error", "issue": f"Evaluator JSON parse failed: {e}"}],
        }
    return verdict


async def run_builder_loop(session: Session, server: PseudoBridgeServer, spec: dict,
                           candidate_dir: Path, results_contract: str = "",
                           model: str = "sonnet") -> None:
    """Long-running Builder loop using a stateful ClaudeSDKClient.

    The client maintains conversation state across turns. After each turn,
    the harness sends "continue" so the builder proceeds. Mid-build feedback
    arrives via fail-files written by the eval watcher (builder_v3.md gates
    on `ls eval_fail_*.json` before the next sheet and before declaring done).
    """
    prompt = (AGENTS_DIR / "builder_v3.md").read_text()

    # Per-session script directory — clear on every builder start to avoid pollution.
    scripts_dir = session.run_dir / "scripts"
    if scripts_dir.exists():
        for old in scripts_dir.glob("*.py"):
            old.unlink()
    scripts_dir.mkdir(exist_ok=True)
    log.info(f"Builder scripts dir: {scripts_dir} (cleaned)")

    # Create a stateful client that persists across turns.
    options = ClaudeAgentOptions(
        system_prompt=prompt,
        allowed_tools=[
            "Read", "Glob", "Grep",
            f"Write({scripts_dir}/builder_*.py)",
            f"Edit({scripts_dir}/builder_*.py)",
            f"Write({candidate_dir}/results.json)",
            f"Edit({candidate_dir}/results.json)",
            f"Bash(python3 {scripts_dir}/builder_*.py)",
            "Bash(ls*)", "Bash(cat*)", f"Bash(rm {session.run_dir}/eval_fail_*.json)",
        ],
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
        model=model,
        max_thinking_tokens=5000,
        add_dirs=[
            str(session.input_dir),
            str(session.run_dir),
        ],
        env={
            "BRIDGE_URL": server.base_url,
            "BRIDGE_VERIFY_TLS": "0",
        },
    )
    client = ClaudeSDKClient(options=options)

    builder_context = build_context_for_builder(session, spec)
    log.info(f"Builder context: {len(builder_context)} chars (~{len(builder_context)//4} tokens)")

    initial_msg = (
        f"The workbook is at {server.base_url} (pseudo-bridge, headless). "
        f"Default Bridge() picks it up via the BRIDGE_URL env var.\n"
        f"Write one script per sheet to {scripts_dir}/builder_<SheetName>.py. Edit in place to fix issues, then re-run. "
        f"IMPORTANT: Set column widths explicitly from the style dumps. Do NOT call auto_fit_columns.\n\n"
        f"On every b.checkpoint(), the workbook auto-saves AND a background evaluator runs. "
        f"Verdicts land at {session.run_dir}/eval_verdicts/checkpoint_*.json. "
        f"On failure, an eval_fail_*.json appears in {session.run_dir}/. Glob for these between sheets "
        f"and before declaring done — fix the relevant script, re-run, delete the fail file.\n\n"
        f"ALL CONTEXT IS BELOW — do NOT use Read to gather text context. "
        f"You may Read screenshot images listed at the bottom for visual reference.\n\n"
        f"{results_contract}"
        f"{builder_context}"
    )

    max_turns = 50
    turn = 0
    checkpoint_count = 0

    # Start the checkpoint watcher — runs for the entire builder session.
    stop_watcher = asyncio.Event()

    eval_tasks = []  # track background eval tasks for cleanup

    async def run_eval_background(cp_num: int, description: str, eval_dir: Path, xlsx_bytes: bytes):
        """Run eval in background — results sent as chat messages."""
        try:
            from timing import mark
            verdict = await run_evaluator(eval_dir, session.input_dir / "dumps")
            mark(f"checkpoint_{cp_num}_eval_done: {verdict.get('status')}")

            if verdict.get("status") == "pass" and not xlsx_bytes:
                verdict = {"status": "fail", "findings": [{"severity": "error", "issue": "Workbook snapshot failed."}]}

            verdicts_dir = session.run_dir / "eval_verdicts"
            verdicts_dir.mkdir(exist_ok=True)
            (verdicts_dir / f"checkpoint_{cp_num}.json").write_text(json.dumps(verdict, indent=2))

            if verdict.get("status") == "pass" and xlsx_bytes:
                model_path = session.run_dir / "models"
                model_path.mkdir(exist_ok=True)
                (model_path / "model.xlsx").write_bytes(xlsx_bytes)
                log.info(f"eval pass: {description}")
            else:
                findings_text = "\n".join(
                    f"  - [{f.get('severity', 'error')}] "
                    f"{f.get('sheet', '')}"
                    f"{(':' + f['cell']) if f.get('cell') else ''}: "
                    f"{f.get('issue', '')}"
                    for f in verdict.get("findings", [])
                )
                log.info(f"eval FAIL: {description}\n{findings_text}")

            # Write fail file — builder globs for eval_fail_*.json between sheets
            # and before declaring done.
            if verdict.get("status") == "fail":
                fail_file = session.run_dir / f"eval_fail_{cp_num}.json"
                fail_file.write_text(json.dumps(verdict, indent=2))
        except Exception as e:
            log.warning(f"checkpoint eval error: {e}")

    async def watch_checkpoints():
        """Poll pseudo_bridge.checkpoint_log for new entries and fire the eval pipeline.

        pseudo_bridge appends to checkpoint_log when the builder calls
        b.checkpoint(). We track our consumption pointer (last_seen) and
        spawn run_eval_background for each new entry.
        """
        nonlocal checkpoint_count
        last_seen = 0
        while not stop_watcher.is_set():
            log_len = len(server.checkpoint_log)
            while last_seen < log_len:
                cp_entry = server.checkpoint_log[last_seen]
                last_seen += 1
                checkpoint_count += 1
                cp_num = checkpoint_count
                description = cp_entry.get("description") or f"checkpoint {cp_num}"
                from timing import mark
                mark(f"checkpoint_{cp_num}_start: {description}")
                eval_dir, xlsx_bytes = await prepare_eval_input(
                    session, server, description, spec, cp_num
                )
                mark(f"checkpoint_{cp_num}_eval_input_ready")
                task = asyncio.create_task(
                    run_eval_background(cp_num, description, eval_dir, xlsx_bytes)
                )
                eval_tasks.append(task)
            await asyncio.sleep(0.3)

    watcher_task = asyncio.create_task(watch_checkpoints())

    try:
        # Connect (starts the subprocess), then send first message
        from timing import mark
        mark("builder_starting")
        log.info("Builder: connecting...")
        await client.connect()
        mark("builder_connected")
        log.info("Builder: connected, sending initial query...")
        await client.query(initial_msg)
        mark("builder_initial_query_sent")

        while turn < max_turns:
            turn += 1
            final_text = ""
            msg_count = 0
            turn_tools = []

            mark(f"builder_turn_{turn}_start")
            async for message in client.receive_messages():
                msg_count += 1
                if isinstance(message, AssistantMessage):
                    for block in message.content:
                        if isinstance(block, TextBlock):
                            final_text = block.text
                            preview = block.text[:100].replace('\n', ' ')
                            log.info(f"Builder text: {preview}...")
                        elif isinstance(block, ToolUseBlock):
                            tool_name = block.name
                            tool_input = block.input or {}
                            turn_tools.append(tool_name)
                            if tool_name in ("Write", "Edit"):
                                path = tool_input.get("file_path", "")
                                short = path.split("/")[-1] if "/" in str(path) else path
                                mark(f"  {tool_name.lower()}:{short}")
                            elif tool_name == "Bash":
                                cmd = str(tool_input.get("command", ""))
                                if "python3" in cmd and "builder" in cmd:
                                    mark(f"  run_script")
                                elif "checkpoint" in cmd.lower():
                                    mark(f"  checkpoint_call")
                                else:
                                    mark(f"  bash:{cmd[:50]}")
                            elif tool_name == "Read":
                                path = str(tool_input.get("file_path", ""))
                                short = path.split("/")[-1] if "/" in path else path
                                mark(f"  read:{short}")
                            else:
                                mark(f"  {tool_name}")
                elif isinstance(message, ResultMessage):
                    tool_summary = ", ".join(f"{t}:{turn_tools.count(t)}" for t in dict.fromkeys(turn_tools))
                    mark(f"builder_turn_{turn}_done ({msg_count} msgs, tools: {tool_summary})")
                    _track_usage("builder", f"turn_{turn}", message)
                    break

            session.append_chat("agent", final_text)

            # Completion sentinel — but enforce the eval-completion gate before
            # accepting it. The builder prompt asks the model to wait for pending
            # evals, but it sometimes declares done eagerly. Without this gate
            # the evaluator's findings land after the builder has exited and
            # the fail-file retry loop is functionally inert.
            if "Model complete. Ready for review." in final_text:
                if eval_tasks:
                    pending = [t for t in eval_tasks if not t.done()]
                    if pending:
                        log.info(f"sentinel received; waiting for {len(pending)} pending eval(s)")
                        await asyncio.gather(*pending, return_exceptions=True)
                fail_files = sorted(session.run_dir.glob("eval_fail_*.json"))
                if fail_files:
                    findings = "\n\n".join(
                        f"## {ff.name}\n{ff.read_text()}" for ff in fail_files
                    )
                    log.info(f"sentinel deferred — {len(fail_files)} fail file(s): "
                             f"{', '.join(f.name for f in fail_files)}")
                    await client.query(
                        f"Pending evaluator failures detected. Address each, re-run "
                        f"the affected sheet, delete the corresponding fail file with "
                        f"`rm`, and re-emit the sentinel only when ALL fail files are "
                        f"gone.\n\n{findings}"
                    )
                    continue
                log.info("builder: completion sentinel verified (no pending fails)")
                return

            # No live user messages in headless — just send "continue" between turns.
            next_prompt = "continue"
            log.info(f"Builder: sending turn {turn + 1} prompt ({len(next_prompt)} chars)...")
            await client.query(next_prompt)

    finally:
        stop_watcher.set()
        await watcher_task
        # Let in-flight evals finish so their verdicts and model.xlsx land
        # before we tear down.
        if eval_tasks:
            log.info(f"waiting for {len(eval_tasks)} in-flight eval(s) to finish")
            await asyncio.gather(*eval_tasks, return_exceptions=True)
        try:
            await client.disconnect()
        except Exception:
            pass

    log.warning(f"builder loop hit max turns ({max_turns})")


def _results_contract_block(task_meta: dict, candidate_dir: Path) -> str:
    """Build the results.json contract section for the Builder prompt.

    If task.yaml defines output_keys, the Builder MUST produce a
    results.json next to model.xlsx mapping each key to a Sheet!Cell
    address. Ported from benchmarks/headless_builder.py to keep the
    headless harness self-contained.
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
    lines.append("Do NOT write this file until all sheets are built and saved.")
    lines.append("")
    return "\n\n---\n\n".join(["\n".join(lines), ""])


async def run_session(*, task_dir: Path, run_dir: Path,
                      model: str = "sonnet",
                      max_turns: int = 50,
                      time_budget_seconds: float | None = None,
                      port: int = 3100) -> dict:
    """Run the full pipeline (planner P1+oracle, P2, P3, builder ↔ evaluator
    loop, grade) on one task. Headless — no live UI, no add-in.

    Returns a result dict with grading + usage breakdown matching the
    contract that benchmarks.experiments.loss + eval_runner expect; also
    writes `run_dir/result.json`.
    """
    import shutil
    import time as _time
    import yaml

    t_start = _time.time()
    run_dir = Path(run_dir)
    run_dir.mkdir(parents=True, exist_ok=True)

    # Session lays out run_dir/{input,scripts,snapshots,...} for us.
    session = Session(root=ROOT, run_dir=run_dir)
    log.info(f"session: {session.run_dir}")

    # --- Read task.yaml early so we know about stub_file before staging ---
    task_yaml_path = task_dir / "task.yaml"
    task_meta = yaml.safe_load(task_yaml_path.read_text()) if task_yaml_path.exists() else {}

    # --- Stage inputs from task_dir/inputs into session.input_dir ---
    task_inputs = task_dir / "inputs"
    if task_inputs.exists():
        for src in task_inputs.iterdir():
            dst = session.input_dir / src.name
            if src.is_dir():
                shutil.copytree(src, dst, dirs_exist_ok=True)
            else:
                shutil.copy2(src, dst)

    # --- Stage the stub xlsx (if any) into session.input_dir so dump.py picks it up.
    # The pseudo-bridge starts from a blank workbook; the builder loads the stub via
    # bridge.copy_sheet_from_input(stub_abs_path, ...) using the absolute path the
    # builder context surfaces.
    staged_stub: Path | None = None
    stub_rel = task_meta.get("stub_file")
    if stub_rel:
        stub_src = task_dir / stub_rel
        if stub_src.exists():
            staged_stub = session.input_dir / stub_src.name
            shutil.copy2(stub_src, staged_stub)

    # --- Brief + clarifications ---
    brief_path = task_dir / "brief.md"
    brief = brief_path.read_text() if brief_path.exists() else ""
    session.save_brief(brief)

    # --- Canonical preprocessing: PDF extraction + xlsx dumping (all sheets) ---
    extract_pdf_inputs(session)
    preprocess_inputs(session)  # no sheet_filter → dump everything (incl. stub if staged)

    # --- Oracle config (simulated user answering planner clarifications) ---
    context_path = task_dir / "context.md"
    oracle_cfg = OracleConfig(
        task_id=task_dir.name,
        context_md=context_path.read_text() if context_path.exists() else "",
        seed_path=task_dir / "clarifications.seed.yaml",
    )

    # --- Task metadata for results contract ---
    tier = task_meta.get("tier")
    cost_budget = task_meta.get("budget", {}).get("cost_dollars") if isinstance(task_meta.get("budget"), dict) else None

    candidate_dir = session.run_dir / "candidate"
    candidate_dir.mkdir(parents=True, exist_ok=True)
    candidate_xlsx = candidate_dir / "model.xlsx"
    results_contract = _results_contract_block(task_meta, candidate_dir)

    # If a stub was staged, prepend a directive telling the builder it's the
    # starting state to copy in (not just a reference). The dump itself is
    # surfaced via build_context_for_builder.
    stub_directive = ""
    if staged_stub is not None:
        stub_directive = (
            f"## Starting state\n"
            f"The workbook starts BLANK. The file `{staged_stub.absolute()}` is the "
            f"stub — it contains pre-filled inputs the brief assumes are present. "
            f"As your FIRST step in the build, use `b.copy_sheet_from_input(...)` "
            f"to load each of its sheets into the candidate workbook. Then proceed "
            f"with the task.\n\n"
        )

    # --- Planner context (brief + dumps + clarification stubs) ---
    planner_context = build_context_for_planner(session, brief)

    # --- Run planner with oracle, then run builder ↔ evaluator loop ---
    spec: dict | None = None
    oracle_stats: dict = {"seed_hits": 0, "cache_hits": 0, "llm_calls": 0,
                          "over_budget": 0, "tokens_in": 0, "tokens_out": 0,
                          "cost_dollars": 0.0}
    completed = False
    terminated_reason = "unknown"
    grading_err: str | None = None
    try:
        with PseudoBridgeServer(output_xlsx=candidate_xlsx, port=port, host="127.0.0.1") as server:
            spec, oracle_stats = await run_planner(session, oracle_cfg, brief,
                                                   planner_context=planner_context)
            log.info(f"spec complete: {len(spec.get('sheets', []))} sheet(s)")

            await run_builder_loop(session, server, spec, candidate_dir,
                                   results_contract=stub_directive + results_contract,
                                   model=model)
            log.info("builder loop done")
            # run_builder_loop returns on completion sentinel or max_turns;
            # treat any return without exception as "completed enough to grade".
            completed = True
            terminated_reason = "sentinel_or_max_turns"
    except Exception as e:
        log.error(f"run_session pipeline error: {type(e).__name__}: {e}")
        terminated_reason = f"crashed:{type(e).__name__}"

    wall = _time.time() - t_start
    over_budget = (time_budget_seconds is not None and wall > time_budget_seconds)
    if over_budget and terminated_reason in ("unknown", "sentinel_or_max_turns"):
        terminated_reason = "time_budget_exceeded"

    # --- Grade ---
    from benchmarks.grader import grade
    graded: dict = {"accuracy": 0.0, "passed": 0, "total": 0,
                    "weighted_score": 0.0, "total_weight": 0.0, "checks": []}
    if candidate_xlsx.exists():
        try:
            graded = grade(
                candidate_xlsx=candidate_xlsx,
                grading_yaml=task_dir / "gold" / "grading.yaml",
                task_dir=task_dir,
            )
        except Exception as e:
            grading_err = f"{type(e).__name__}: {e}"
            log.error(f"grader failed: {grading_err}")

    # --- Aggregate session usage into per-agent + dollar totals ---
    builder_usage = _SESSION_USAGE.get("builder", {})
    planner_usage = _SESSION_USAGE.get("planner", {})
    evaluator_usage = _SESSION_USAGE.get("evaluator", {})
    dollars_total = sum(b.get("cost_usd", 0.0) for b in _SESSION_USAGE.values())

    # planner_stats has historically been a flat dict with prefixed keys
    # consumed by loss.cold_cost_from_result. Mirror that shape so the
    # cold-cost accounting in loss.py picks up planner tokens.
    planner_stats_flat = {
        "planner_model": model,
        "planner_input_tokens": planner_usage.get("input", 0),
        "planner_output_tokens": planner_usage.get("output", 0),
        "planner_cache_creation_input_tokens": planner_usage.get("cache_w", 0),
        "planner_cache_read_input_tokens": planner_usage.get("cache_r", 0),
        "planner_cost_usd": planner_usage.get("cost_usd", 0.0),
        "planner_calls": planner_usage.get("calls", 0),
    }
    evaluator_stats_flat = {
        "calls": evaluator_usage.get("calls", 0),
        "cost_usd": evaluator_usage.get("cost_usd", 0.0),
        "input_tokens": evaluator_usage.get("input", 0),
        "output_tokens": evaluator_usage.get("output", 0),
    }

    result = {
        "task_id": task_dir.name,
        "tier": tier,
        "cost_budget_dollars": cost_budget,
        "run_dir": str(session.run_dir),
        "candidate": str(candidate_xlsx),
        "spec_generated": spec is not None,
        "spec_sheets": len(spec.get("sheets", [])) if spec else 0,
        "completed": completed,
        "terminated_reason": terminated_reason,
        "wall_seconds": wall,
        "time_budget_seconds": time_budget_seconds,
        "over_budget": over_budget,
        "builder_model": model,
        "builder_usage": builder_usage,
        "planner_stats": planner_stats_flat,
        "evaluator_stats": evaluator_stats_flat,
        "oracle_stats": oracle_stats,
        "dollars": dollars_total,
        "accuracy": graded.get("accuracy", 0.0),
        "passed": graded.get("passed", 0),
        "total": graded.get("total", 0),
        "weighted_score": graded.get("weighted_score", 0.0),
        "total_weight": graded.get("total_weight", 0.0),
        "checks": graded.get("checks", []),
        "session_usage": _SESSION_USAGE,
    }
    (session.run_dir / "result.json").write_text(json.dumps(result, indent=2, default=str))
    _log_session_totals()
    return result


async def main() -> None:
    """CLI entrypoint. Usage: python3 -m harness --task <task_id> [--run-dir DIR]"""
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--task", required=True, help="task_id under benchmarks/tasks/")
    ap.add_argument("--run-dir", default=None,
                    help="Output dir; default: benchmarks/runs/<timestamp>_<task>")
    ap.add_argument("--model", default="sonnet")
    ap.add_argument("--max-turns", type=int, default=50)
    ap.add_argument("--port", type=int, default=3100)
    args = ap.parse_args()

    task_dir = ROOT / "benchmarks" / "tasks" / args.task
    if not task_dir.exists():
        raise SystemExit(f"task not found: {task_dir}")

    if args.run_dir:
        run_dir = Path(args.run_dir)
    else:
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        run_dir = ROOT / "benchmarks" / "runs" / f"{ts}_{args.task}"

    result = await run_session(
        task_dir=task_dir, run_dir=run_dir,
        model=args.model, max_turns=args.max_turns, port=args.port,
    )
    print(json.dumps({"task": args.task, "accuracy": result["accuracy"],
                      "passed": result["passed"], "total": result["total"],
                      "run_dir": result["run_dir"]}, indent=2))


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("Stopped")
