#!/usr/bin/env python3
"""ExcelHarness v3 — single-process live-Excel harness.

Flow:
  1. Start BridgeServer (HTTPS + WSS).
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

from bridge_server import BridgeServer
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
CERTS = ROOT / "officejs-prototype" / "certs"
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

    # Include formula dumps
    for dump_dir in sorted((session.input_dir / "dumps").iterdir()):
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

    # Deduplicated styles — one per sheet type
    dump_dirs = sorted((session.input_dir / "dumps").iterdir())
    for f in _dedupe_styles(dump_dirs):
        parts.append(f"## Style: {f.name}\n{f.read_text()}")

    # All formula text dumps — builder needs data sources + reference formulas
    for dump_dir in dump_dirs:
        formulas_dir = dump_dir / "formulas"
        if formulas_dir.exists():
            for f in sorted(formulas_dir.glob("*.txt")):
                parts.append(f"## Formulas: {f.name}\n{f.read_text()}")

    # List screenshot paths so builder can Read them for visual reference
    screenshots = []
    for dump_dir in dump_dirs:
        ss_dir = dump_dir / "screenshots"
        if ss_dir.exists():
            for f in sorted(ss_dir.glob("*.jpg")) + sorted(ss_dir.glob("*.png")):
                screenshots.append(str(f))
    if screenshots:
        parts.append("## Reference Screenshots (use Read to view)\n" + "\n".join(screenshots))

    return "\n\n---\n\n".join(parts)


async def wait_for_addin(server: BridgeServer, timeout: float = 300.0) -> None:
    """Block until the add-in connects."""
    deadline = asyncio.get_event_loop().time() + timeout
    while server._addin_ws is None:
        if asyncio.get_event_loop().time() > deadline:
            raise RuntimeError("Timed out waiting for add-in to connect")
        await asyncio.sleep(0.2)


async def ask_user(server: BridgeServer, prompt: str) -> str:
    """Send a prompt to the taskpane chat and wait for the user's reply."""
    await server.send_chat(prompt)
    while True:
        pending = server.chat_queue.drain_all()
        if pending:
            return pending[-1]
        await asyncio.sleep(0.3)


async def run_planner(session: Session, server: BridgeServer, brief: str, planner_context: str | None = None) -> dict:
    """Planner using ClaudeSDKClient — single session, all passes share context."""
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
        await server.send_chat("Analyzing your brief for ambiguities...")

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

        # --- Ask user each question via chat, with inline pruning by the Planner ---
        # After each answer we ask the Planner (same stateful session, context
        # cached) which remaining questions are now resolved and which still
        # need asking. This avoids redundant questions without adding a new
        # agent — the Planner already has the full spec context.
        clarifications: dict[str, str] = {}
        remaining: list[dict] = list(questions)
        pruned_count = 0

        while remaining:
            q = remaining.pop(0)
            qid = q.get("id", "unknown")
            question_text = q.get("question", "")
            choices = q.get("choices", [])

            msg = f"**{question_text}**"
            if q.get("context"):
                msg += f"\n\n_{q['context']}_"
            if choices:
                msg += "\n\nOptions:\n" + "\n".join(f"  {i+1}. {c}" for i, c in enumerate(choices))

            answer = await ask_user(server, msg)
            clarifications[qid] = answer
            session.append_clarification(qid, answer)
            session.append_chat("agent", msg)
            session.append_chat("user", answer)

            # Prune remaining questions based on the latest answer.
            if remaining:
                pruned_count += 1
                prune_msg = (
                    f"The user just answered `{qid}` with: {answer}\n\n"
                    f"Here are the STILL-PENDING questions from Pass 1, as a JSON array:\n"
                    f"{json.dumps(remaining, indent=2)}\n\n"
                    "Given this answer (and all prior answers), return a JSON object "
                    "indicating which pending questions are now resolved by strong implication, "
                    "and which still genuinely need to be asked:\n"
                    "```json\n"
                    "{\n"
                    '  "skip": [{"id": "question_id", "derived_answer": "concise answer"}],\n'
                    '  "ask": ["question_id_still_needed", "..."]\n'
                    "}\n"
                    "```\n"
                    "Every pending id must appear in exactly one of `skip` or `ask`. "
                    "When in doubt, put it in `ask`. Respond with ONLY the JSON object."
                )
                prune_text = await get_response(f"prune_{pruned_count}", prune_msg)
                pm = re.search(r"\{.*\}", prune_text, re.DOTALL)
                if pm:
                    try:
                        data = json.loads(pm.group(0))
                        skip = data.get("skip", []) or []
                        ask = set(data.get("ask", []) or [])
                        skipped_ids = {s.get("id"): s.get("derived_answer", "") for s in skip if s.get("id")}
                        if skipped_ids:
                            for sid, derived in skipped_ids.items():
                                clarifications[sid] = f"(auto-derived from prior answer) {derived}"
                                session.append_clarification(sid, clarifications[sid])
                            skipped_list = ", ".join(skipped_ids.keys())
                            log.info(f"Planner pruned {len(skipped_ids)} question(s): {skipped_list}")
                            await server.send_chat(
                                f"Skipping {len(skipped_ids)} question(s) resolved by your last answer: {skipped_list}"
                            )
                        # Keep only questions still flagged as `ask` (and not in `skip`).
                        remaining = [
                            rq for rq in remaining
                            if rq.get("id") in ask and rq.get("id") not in skipped_ids
                        ]
                    except json.JSONDecodeError:
                        log.warning(f"Prune pass {pruned_count} returned unparsable JSON; asking all remaining.")

        # --- Pass 2: Full spec (same session — context cached) ---
        mark(f"qa_complete ({len(clarifications)} answers)")
        await server.send_chat("Generating spec...")
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
                await server.send_chat("Spec extracted from response (Write tool didn't fire).")
            else:
                log.info(f"Pass 2 text preview: {pass2_text[:500]}")
                raise RuntimeError("Planner did not write model_spec.json and no JSON found in response")

        spec = json.loads(spec_path.read_text())
        try:
            jsonschema.validate(instance=spec, schema=schema)
        except jsonschema.ValidationError as e:
            await server.send_chat(f"Spec failed validation: {e.message}. Fixing...")
            fix_text = await get_response("pass2_fix", f"Your spec failed validation: {e.message}\n\nFix it and save again.")
            spec = json.loads(spec_path.read_text())
            jsonschema.validate(instance=spec, schema=schema)

        # --- Pass 3: Self-review (same session — has full context) ---
        mark("pass2_validated")
        await server.send_chat("Reviewing spec...")
        pass3_text = await get_response(
            "pass3",
            "Run Pass 3 (self-review). Check for contradictions, overprescription, "
            "incompleteness, or redundancy. If you find issues, rewrite the spec. "
            "If it's clean, reply 'CLEAN'."
        )

        # Reload in case pass 3 rewrote
        spec = json.loads(spec_path.read_text())
        jsonschema.validate(instance=spec, schema=schema)

        return spec

    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


async def prepare_eval_input(
    session: Session,
    server: BridgeServer,
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

    # Snapshot the workbook. During shutdown the add-in can return a partial
    # response (ok without base64), so gate on both keys before decoding.
    snap_result = await server.send_command("saveSnapshot", {}, timeout=180.0)
    xlsx_bytes = b""
    snap_path: Path | None = None
    if snap_result.get("ok") and snap_result.get("base64"):
        xlsx_bytes = base64.b64decode(snap_result["base64"])
        snap_path = session.snapshots_dir / f"turn_{turn}.xlsx"
        snap_path.write_bytes(xlsx_bytes)
    elif snap_result.get("ok"):
        log.warning("saveSnapshot returned ok with no base64 payload — treating as snapshot failure.")

        # Render PNGs
        try:
            render_xlsx_to_pngs(snap_path, eval_dir / "screenshots")
        except Exception as e:
            # Rendering failure is not fatal — evaluator still has dumps + spec.
            log.warning(f"Snapshot render failed: {e}")

    # Dump each sheet from the snapshot xlsx via openpyxl — avoids Office.js
    # recalc + RPC timeouts that previously capped sheets at 30s.
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


async def run_builder_loop(session: Session, server: BridgeServer, spec: dict) -> None:
    """Long-running Builder loop using a stateful ClaudeSDKClient.

    The client maintains conversation state across turns, so the Builder
    remembers what it built. Between turns, the harness drains the chat
    queue and injects user messages as the next prompt.
    """
    prompt = (AGENTS_DIR / "builder_v3.md").read_text()

    # Clear all existing sheets so we start fresh (keep one sheet — Excel requires at least one).
    try:
        names_resp = await server.send_command("getSheetNames", {})
        existing = names_resp.get("sheets", [])
        if existing:
            # Create a temp sheet, delete all others, then builder will replace it
            await server.send_command("createSheet", {"name": "__temp__"})
            for name in existing:
                try:
                    await server.send_command("deleteSheet", {"name": name})
                except Exception:
                    pass
            log.info(f"Cleared {len(existing)} existing sheet(s) for fresh build.")
    except Exception as e:
        log.warning(f"could not clear sheets: {e}")

    # Protect the workbook so user can't edit during build (non-fatal if it fails).
    protect_result = await server.send_command("protectWorkbook", {})
    if protect_result.get("error"):
        log.warning(f"Workbook protection failed (non-fatal): {protect_result['error']}")

    # Per-session script directory — clear on every builder start to avoid pollution
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
            f"Bash(python3 {scripts_dir}/builder_*.py)",
            "Bash(ls*)", "Bash(cat*)", f"Bash(rm {session.run_dir}/eval_fail_*.json)",
        ],
        permission_mode="bypassPermissions",
        cwd=str(ROOT),
        model="sonnet",  # Sonnet builder — faster + cheaper than Opus; accuracy validated on prior runs
        max_thinking_tokens=5000,  # Cap adaptive thinking — without this, Sonnet can burn its full 64K thinking budget (~24 min at 44 t/s) before emitting any output. 5K is enough for tool-call reasoning but bounds the silent stall.
        add_dirs=[
            str(session.input_dir),   # input xlsx + dumps (styles, formulas, screenshots)
            str(session.run_dir),     # spec, chat log, scripts
        ],
    )
    client = ClaudeSDKClient(options=options)

    # Inject everything — spec + bridge + all styles + data source formulas + screenshot paths
    builder_context = build_context_for_builder(session, spec)
    log.info(f"Builder context: {len(builder_context)} chars (~{len(builder_context)//4} tokens)")

    initial_msg = (
        f"The workbook is live at https://localhost:3000. "
        f"Write one script per sheet to {scripts_dir}/builder_<SheetName>.py. Edit in place to fix issues, then re-run. "
        f"IMPORTANT: Set column widths explicitly from the style dumps. Do NOT call auto_fit_columns.\n\n"
        f"ALL CONTEXT IS BELOW — do NOT use Read to gather text context. "
        f"You may Read screenshot images listed at the bottom for visual reference.\n\n"
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
                # Overwrite the latest-known-good model file each time a checkpoint
                # passes. Not committed to git — the per-turn snapshots in
                # session.snapshots_dir already provide recovery history.
                model_path = session.run_dir / "models"
                model_path.mkdir(exist_ok=True)
                (model_path / "model.xlsx").write_bytes(xlsx_bytes)
                await server.send_chat(f"✓ Eval passed: {description}")
            else:
                findings_text = "\n".join(
                    f"  - [{f.get('severity', 'error')}] "
                    f"{f.get('sheet', '')}"
                    f"{(':' + f['cell']) if f.get('cell') else ''}: "
                    f"{f.get('issue', '')}"
                    for f in verdict.get("findings", [])
                )
                await server.send_chat(f"✗ Eval FAIL: {description}\n{findings_text}")

            # Write verdict to per-checkpoint file the builder can check
            # Only write fail files — builder globs for any eval_fail_*.json
            if verdict.get("status") == "fail":
                fail_file = session.run_dir / f"eval_fail_{cp_num}.json"
                fail_file.write_text(json.dumps(verdict, indent=2))
        except Exception as e:
            try:
                await server.send_chat(f"Checkpoint eval error: {e}")
            except Exception:
                pass

    async def watch_checkpoints():
        nonlocal checkpoint_count
        while not stop_watcher.is_set():
            pending_cp = server.pop_pending_checkpoint()
            if pending_cp is not None:
                checkpoint_count += 1
                cp_num = checkpoint_count
                from timing import mark
                mark(f"checkpoint_{cp_num}_start: {pending_cp.description}")
                eval_dir, xlsx_bytes = await prepare_eval_input(
                    session, server, pending_cp.description, spec, cp_num
                )
                mark(f"checkpoint_{cp_num}_eval_input_ready")

                # Resolve immediately — builder continues
                pending_cp.resolve({"status": "pass", "findings": [], "_pending_eval": True})
                await server.send_chat(f"⏳ Evaluating: {pending_cp.description}...")

                # Spawn eval as independent task — multiple evals can run in parallel
                task = asyncio.create_task(
                    run_eval_background(cp_num, pending_cp.description, eval_dir, xlsx_bytes)
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
                                await server.send_chat(f"{'Writing' if tool_name == 'Write' else 'Editing'} {short}...")
                            elif tool_name == "Bash":
                                cmd = str(tool_input.get("command", ""))
                                if "python3" in cmd and "builder" in cmd:
                                    mark(f"  run_script")
                                    await server.send_chat(f"Running build script...")
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

            # Check for completion sentinel
            if "Model complete. Ready for review." in final_text:
                await server.send_command("unprotectWorkbook", {})
                await server.send_chat("Session complete. Workbook unprotected for manual edits.")
                return

            # Drain chat for next turn
            user_msgs = server.chat_queue.drain_all()
            if user_msgs:
                for m in user_msgs:
                    session.append_chat("user", m)
                next_prompt = "\n\n".join(user_msgs)
            else:
                next_prompt = "continue"

            # Send next turn to the stateful client
            log.info(f"Builder: sending turn {turn + 1} prompt ({len(next_prompt)} chars)...")
            await client.query(next_prompt)

    finally:
        stop_watcher.set()
        await watcher_task
        # Let in-flight evals finish so their verdicts and model.xlsx land
        # before we tear down. Without this, the final checkpoint's eval
        # can be orphaned by loop shutdown.
        if eval_tasks:
            await server.send_chat(f"Waiting for {len(eval_tasks)} in-flight eval(s) to finish...")
            await asyncio.gather(*eval_tasks, return_exceptions=True)
        try:
            await client.disconnect()
        except Exception:
            pass

    # Fell out of loop — max turns exceeded
    await server.send_chat(f"Builder loop hit max turns ({max_turns}). Stopping.")
    await server.send_command("unprotectWorkbook", {})


async def main() -> None:
    import sys
    resume_ts = None
    for arg in sys.argv[1:]:
        if arg.startswith("--resume="):
            resume_ts = arg.split("=", 1)[1]
        elif arg == "--resume":
            # Find latest session with a spec
            runs = sorted((ROOT / "runs").glob("*"), reverse=True)
            for r in runs:
                if (r / "model_spec.json").exists():
                    resume_ts = r.name
                    break
    session = Session(root=ROOT, timestamp=resume_ts)
    log.info(f"Session: {session.run_dir}")

    server = BridgeServer(
        host="localhost",
        http_port=3000,
        wss_port=3001,
        cert_path=CERTS / "cert.pem",
        key_path=CERTS / "key.pem",
    )
    await server.start()
    log.info("Bridge server listening on :3000/:3001")
    log.info("Open Excel with the sideloaded add-in now.")

    try:
        await wait_for_addin(server)
        log.info("Add-in connected.")

        # Check if spec already exists (resume from Builder phase)
        spec_path = session.run_dir / "model_spec.json"
        log.info(f"Checking for spec at {spec_path} ... exists={spec_path.exists()}")
        if spec_path.exists():
            spec = json.loads(spec_path.read_text())
            log.info(f"Found existing spec with {len(spec['sheets'])} sheets. Sending chat...")
            await server.send_chat(f"Resuming with existing spec ({len(spec['sheets'])} sheets). Starting Builder...")
            log.info(f"Chat sent. Jumping to Builder.")
        else:
            from timing import mark, reset
            reset()
            mark("waiting_for_brief")
            brief = await ask_user(server, f"Hi. What would you like to build? Drop input files into {session.input_dir} first, then paste your brief.")
            mark("brief_received")
            session.save_brief(brief)
            session.append_chat("user", brief)

            # Extract any PDFs to companion .txt so the planner context
            # (which globs *.txt) picks them up automatically.
            extract_pdf_inputs(session)

            # Sheet mapper — user picks which sheets to dump
            import openpyxl as _xl
            xlsx_files = sorted(session.input_dir.glob("*.xlsx")) + sorted(session.input_dir.glob("*.xls"))
            if xlsx_files:
                mapper_files = []
                for xlsx in xlsx_files:
                    try:
                        wb = _xl.load_workbook(xlsx, read_only=True)
                        mapper_files.append({"name": xlsx.name, "sheets": wb.sheetnames})
                        wb.close()
                    except Exception as e:
                        log.warning(f"couldn't read {xlsx.name}: {e}")
                if mapper_files:
                    mark("sheet_mapper_sent")
                    selected = await server.send_sheet_mapper(mapper_files)
                    mark(f"sheet_mapper_done ({sum(len(v) for v in selected.values())} sheets selected)")
                    # Build sheet filter: map filename -> list of sheet names
                    sheet_filter = {}
                    for xlsx in xlsx_files:
                        if xlsx.name in selected:
                            sheet_filter[xlsx] = selected[xlsx.name]
                    xlsx_files = preprocess_inputs(session, sheet_filter=sheet_filter)
                else:
                    xlsx_files = preprocess_inputs(session)
            else:
                xlsx_files = []
            mark("dumps_complete")
            if xlsx_files:
                await server.send_chat(f"Preprocessed {len(xlsx_files)} input file(s). Starting planning...")
            else:
                await server.send_chat("No input xlsx files found. Starting planning...")

            # Build consolidated context and inject into Planner prompt
            planner_context = build_context_for_planner(session, brief)
            mark(f"planner_context_built ({len(planner_context)//4} tokens)")

            spec = await run_planner(session, server, brief, planner_context=planner_context)
            mark("spec_complete")
            await server.send_chat(f"Spec complete: {len(spec['sheets'])} sheet(s). Saved to {session.run_dir.name}/model_spec.json")
            log.info(f"Spec saved. Planner phase done.")

        await run_builder_loop(session, server, spec)
        log.info(f"Builder loop done.")
    finally:
        try:
            await server.send_command("unprotectWorkbook", {})
        except Exception:
            pass
        await server.stop()
        _log_session_totals()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        log.info("Stopped")
