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
import re
import signal
import sys
from pathlib import Path

from bridge_server import BridgeServer
from session import Session

from claude_agent_sdk import query
from claude_agent_sdk.types import (
    AssistantMessage,
    ClaudeAgentOptions,
    ResultMessage,
    TextBlock,
    ToolResultBlock,
    ToolUseBlock,
)
import jsonschema

ROOT = Path(__file__).parent
CERTS = ROOT / "officejs-prototype" / "certs"
AGENTS_DIR = ROOT / "agents"


async def run_agent_session(
    system_prompt: str,
    user_messages: list[str],
    allowed_tools: list[str],
    cwd: Path,
    on_activity: callable | None = None,
) -> str:
    """Run a single-turn Claude Agent SDK query and return the text of the final assistant response.

    on_activity: optional async callback(str) called with status updates as the
    agent works (tool calls, progress). Used to stream updates to chat so the
    user isn't staring at a blank screen.
    """
    options = ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=allowed_tools,
        permission_mode="acceptEdits",
        cwd=str(cwd),
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
                        if "python3 /tmp/builder" in cmd:
                            await on_activity(f"Running build script...")
    return final_text


def preprocess_inputs(session: Session, screenshots: bool = False) -> list[Path]:
    """Run dump.py on each xlsx in the session's input directory.

    Produces text dumps (formulas, values, styles) into input/dumps/.
    Optionally renders screenshots (slow — only needed for Builder, not Planner).
    Returns list of xlsx files found.
    """
    import subprocess
    input_dir = session.input_dir
    xlsx_files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xls"))
    if not xlsx_files:
        return []

    dumps_dir = input_dir / "dumps"
    dumps_dir.mkdir(exist_ok=True)

    # dump.py writes to ROOT/evals/ by default. We'll run it, then move results
    # into the session's input/dumps/ directory.
    evals_dir = ROOT / "evals"
    for xlsx in xlsx_files:
        print(f"[harness] Preprocessing {xlsx.name}...")
        # Clear evals/ so we get only this file's output
        import shutil
        if evals_dir.exists():
            shutil.rmtree(evals_dir)

        # Run dump.py — it writes to ROOT/evals/
        result = subprocess.run(
            [sys.executable, str(ROOT / "dump.py"), str(xlsx)],
            capture_output=True, text=True, cwd=str(ROOT),
        )
        if result.returncode != 0:
            print(f"[harness] Warning: dump.py failed on {xlsx.name}: {result.stderr[:200]}")
            continue

        # Move text dumps into input/dumps/<filename>/
        file_dumps = dumps_dir / xlsx.stem
        file_dumps.mkdir(exist_ok=True)
        for subdir in ["formulas", "formulas_json", "styles", "values"]:
            src = evals_dir / subdir
            if src.exists():
                dst = file_dumps / subdir
                if dst.exists():
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)

        # Screenshots only if requested (slow, Planner doesn't need them)
        if screenshots:
            src = evals_dir / "screenshots"
            if src.exists():
                dst = file_dumps / "screenshots"
                if dst.exists():
                    shutil.rmtree(dst)
                shutil.copytree(src, dst)

    # Clean up the global evals/ directory
    import shutil
    if evals_dir.exists():
        shutil.rmtree(evals_dir)

    return xlsx_files


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


async def run_planner(session: Session, server: BridgeServer, brief: str) -> dict:
    """Three-pass Planner: ambiguities -> spec -> self-review."""
    prompt = (AGENTS_DIR / "planner_v3.md").read_text()
    schema = json.loads((ROOT / "schemas" / "model_spec.schema.json").read_text())

    async def status(msg: str):
        await server.send_chat(msg)

    # --- Pass 1: Ambiguities ---
    await server.send_chat("Analyzing your brief for ambiguities...")

    # Tell the Planner exactly where to look — prevents it from wandering into old runs
    session_context = (
        f"SESSION DIRECTORY: {session.run_dir}\n"
        f"INPUT FILES: {session.input_dir}\n"
        f"INPUT DUMPS (text extracts of xlsx): {session.input_dir / 'dumps'}\n"
        f"CONVENTIONS: {ROOT / 'conventions.md'}\n"
        f"SCHEMA: {ROOT / 'schemas' / 'model_spec.schema.json'}\n"
        f"IMPORTANT: Only read files from the paths listed above. Do NOT read from other runs/ directories.\n"
    )

    pass1_user = (
        f"{session_context}\n"
        f"BRIEF:\n{brief}\n\n"
        "Run Pass 1 (ambiguity detection). Output a JSON array of questions, "
        "or an empty array if the brief is fully specified. Respond with ONLY the JSON array."
    )

    pass1_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[pass1_user],
        allowed_tools=["Read", "Glob", "Grep"],
        cwd=ROOT,
        on_activity=status,
    )

    try:
        questions = json.loads(pass1_text.strip())
    except json.JSONDecodeError:
        # Try to extract JSON from markdown code block
        m = re.search(r"\[.*\]", pass1_text, re.DOTALL)
        if m:
            questions = json.loads(m.group(0))
        else:
            questions = []

    # --- Ask user each question via chat ---
    clarifications = {}
    for q in questions:
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

    # --- Pass 2: Full spec ---
    await server.send_chat("Generating spec...")
    clarif_text = "\n".join(f"- {k}: {v}" for k, v in clarifications.items())
    pass2_user = (
        f"{session_context}\n"
        f"BRIEF:\n{brief}\n\n"
        f"CLARIFICATIONS:\n{clarif_text or '(none)'}\n\n"
        "Run Pass 2. Write the full spec and save it to "
        f"{session.run_dir / 'model_spec.json'} "
        "using the Write tool. Then output the JSON of the spec in your final message."
    )

    pass2_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[pass2_user],
        allowed_tools=["Read", "Glob", "Grep", f"Write({session.run_dir}/model_spec.json)"],
        cwd=ROOT,
        on_activity=status,
    )

    # --- Load, validate, optionally retry on schema failure ---
    spec_path = session.run_dir / "model_spec.json"
    if not spec_path.exists():
        # Planner returned the spec as text instead of using Write tool — extract it
        m = re.search(r"```json\s*(\{.*?\})\s*```", pass2_text, re.DOTALL)
        if not m:
            m = re.search(r"\{.*\}", pass2_text, re.DOTALL)
        if m:
            extracted = m.group(1) if m.lastindex else m.group(0)
            spec_path.write_text(extracted)
            await server.send_chat("Spec extracted from Planner response (tool didn't fire). Validating...")
        else:
            raise RuntimeError("Planner did not write model_spec.json and no JSON found in response")

    spec = json.loads(spec_path.read_text())
    try:
        jsonschema.validate(instance=spec, schema=schema)
    except jsonschema.ValidationError as e:
        await server.send_chat(f"Spec failed validation: {e.message}. Retrying...")
        retry_user = (
            f"Your spec failed validation: {e.message}\n\n"
            "Fix it and save again."
        )
        await run_agent_session(
            system_prompt=prompt,
            user_messages=[pass2_user, retry_user],
            allowed_tools=["Read", "Glob", "Grep", f"Write({session.run_dir}/model_spec.json)", f"Edit({session.run_dir}/model_spec.json)"],
            cwd=ROOT,
        )
        spec = json.loads(spec_path.read_text())
        try:
            jsonschema.validate(instance=spec, schema=schema)
        except jsonschema.ValidationError as e:
            await server.send_chat(f"Planner failed to produce a valid spec after retry: {e.message}. Aborting.")
            raise

    # --- Pass 3: Self-review ---
    await server.send_chat("Reviewing spec for contradictions...")
    pass3_user = (
        "Run Pass 3 (self-review). Re-read the spec you just wrote at "
        f"{spec_path} and check for contradictions, overprescription, incompleteness, "
        "or redundancy. If you find issues, rewrite the spec. If it's clean, reply 'CLEAN'."
    )
    pass3_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[pass3_user],
        allowed_tools=["Read", "Glob", "Grep", f"Write({session.run_dir}/model_spec.json)", f"Edit({session.run_dir}/model_spec.json)"],
        cwd=ROOT,
        on_activity=status,
    )
    # Reload and revalidate in case pass 3 rewrote.
    spec = json.loads(spec_path.read_text())
    jsonschema.validate(instance=spec, schema=schema)

    return spec


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

    # Snapshot the workbook
    snap_result = await server.send_command("saveSnapshot", {})
    xlsx_bytes = b""
    if snap_result.get("ok"):
        xlsx_bytes = base64.b64decode(snap_result["base64"])
        snap_path = session.snapshots_dir / f"turn_{turn}.xlsx"
        snap_path.write_bytes(xlsx_bytes)

        # Render PNGs
        try:
            render_xlsx_to_pngs(snap_path, eval_dir / "screenshots")
        except Exception as e:
            # Rendering failure is not fatal — evaluator still has dumps + spec.
            print(f"[harness] Snapshot render failed: {e}")

    # Dump each sheet
    sheets_result = await server.send_command("getSheetNames", {})
    for sheet_name in sheets_result.get("sheets", []):
        dump = await server.send_command("dumpSheet", {"sheet": sheet_name})
        safe_name = sheet_name.replace("/", "_").replace(" ", "_")
        (eval_dir / "dumps" / f"{safe_name}.json").write_text(json.dumps(dump, indent=2))

    return eval_dir, xlsx_bytes


async def run_evaluator(session: Session, eval_dir: Path) -> dict:
    """Invoke the Evaluator agent in a fresh context and parse its JSON verdict."""
    prompt = (AGENTS_DIR / "evaluator_v3.md").read_text()
    user_msg = (
        f"Review the checkpoint. Input is at {eval_dir}. "
        f"Return ONLY a JSON object per your output format spec."
    )

    result_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[user_msg],
        allowed_tools=["Read", "Glob", "Grep", "Bash(cat*)", "Bash(ls*)"],
        cwd=ROOT,
    )

    # Extract JSON from response — try fenced JSON block first
    m = re.search(r"```json\s*(\{.*?\})\s*```", result_text, re.DOTALL)
    if not m:
        # Fallback: find outermost {...}
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
    """Long-running Builder loop.

    Delivers turns to a Claude Agent SDK client session; between turns, drains
    the chat queue and injects as the next user turn. Handles checkpoints by
    snapshotting + (stubbed) evaluator + git commit.
    """
    prompt = (AGENTS_DIR / "builder_v3.md").read_text()

    # Protect the workbook so user can't edit during build.
    await server.send_command("protectWorkbook", {})

    initial_msg = (
        f"Begin building per the spec at {session.run_dir}/model_spec.json. "
        f"The workbook is live at https://localhost:3000. Read the spec, "
        f"then conventions.md, then start."
    )

    pending_messages = [initial_msg]
    max_turns = 50
    turn = 0

    while turn < max_turns:
        turn += 1
        user_msg = "\n\n---\n\n".join(pending_messages)
        pending_messages = []

        # Launch a checkpoint watcher task that runs in parallel with the agent turn.
        stop_watcher = asyncio.Event()
        checkpoint_results: list = []

        async def watch_checkpoints():
            import subprocess
            while not stop_watcher.is_set():
                pending_cp = server.pop_pending_checkpoint()
                if pending_cp is not None:
                    try:
                        # Prepare eval input (snapshot + dumps + screenshots)
                        eval_dir, xlsx_bytes = await prepare_eval_input(
                            session, server, pending_cp.description, spec, turn
                        )

                        # Run the real Evaluator
                        verdict = await run_evaluator(session, eval_dir)

                        # C2: Override pass verdict if snapshot failed
                        if verdict.get("status") == "pass" and not xlsx_bytes:
                            verdict = {"status": "fail", "findings": [{"severity": "error", "issue": "Workbook snapshot failed — cannot commit checkpoint."}]}

                        # M1: Persist evaluator verdict
                        verdicts_dir = session.run_dir / "eval_verdicts"
                        verdicts_dir.mkdir(exist_ok=True)
                        (verdicts_dir / f"turn_{turn}.json").write_text(json.dumps(verdict, indent=2))

                        # Only commit on PASS
                        if verdict.get("status") == "pass" and xlsx_bytes:
                            model_path = session.run_dir / "models"
                            model_path.mkdir(exist_ok=True)
                            committed = model_path / "model.xlsx"
                            committed.write_bytes(xlsx_bytes)
                            add_result = subprocess.run(["git", "add", str(committed)], capture_output=True, text=True, cwd=ROOT)
                            if add_result.returncode != 0:
                                await server.send_chat(f"Warning: git add failed: {add_result.stderr.strip()}")
                            else:
                                commit_result = subprocess.run(
                                    ["git", "commit", "-m", f"Checkpoint: {pending_cp.description}"],
                                    capture_output=True, text=True,
                                    cwd=ROOT,
                                )
                                if commit_result.returncode != 0:
                                    await server.send_chat(f"Warning: git commit failed: {commit_result.stderr.strip()}")
                            await server.send_chat(f"✓ {pending_cp.description}")
                        else:
                            # FAIL (or snapshot failed): don't commit. Pass findings back to Builder.
                            findings_text = "\n".join(
                                f"  - [{f.get('severity', 'error')}] "
                                f"{f.get('sheet', '')}"
                                f"{(':' + f['cell']) if f.get('cell') else ''}: "
                                f"{f.get('issue', '')}"
                                for f in verdict.get("findings", [])
                            )
                            await server.send_chat(
                                f"✗ Evaluator FAIL: {pending_cp.description}\n{findings_text}"
                            )

                        pending_cp.resolve(verdict)
                        checkpoint_results.append(pending_cp.description)
                    except Exception as e:
                        if pending_cp is not None:
                            pending_cp.resolve({"status": "fail", "findings": [{"severity": "error", "issue": f"Checkpoint handler crashed: {e}"}]})
                        try:
                            await server.send_chat(f"Checkpoint handler error: {e}")
                        except Exception:
                            pass
                await asyncio.sleep(0.3)

        watcher_task = asyncio.create_task(watch_checkpoints())

        async def builder_status(msg: str):
            await server.send_chat(msg)

        try:
            final_text = await run_agent_session(
                system_prompt=prompt,
                user_messages=[user_msg],
                allowed_tools=[
                    "Read", "Glob", "Grep",
                    "Write(/tmp/builder_*.py)",
                    "Bash(python3 /tmp/builder_*.py)",
                    "Bash(ls*)", "Bash(cat*)",
                ],
                cwd=ROOT,
                on_activity=builder_status,
            )
        finally:
            stop_watcher.set()
            await watcher_task

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
            pending_messages.extend(user_msgs)
        else:
            pending_messages.append("continue")

    # Fell out of loop — max turns exceeded
    await server.send_chat(f"Builder loop hit max turns ({max_turns}). Stopping.")
    await server.send_command("unprotectWorkbook", {})


async def main() -> None:
    session = Session(root=ROOT)
    print(f"[harness] Session: {session.run_dir}")

    server = BridgeServer(
        host="localhost",
        http_port=3000,
        wss_port=3001,
        cert_path=CERTS / "cert.pem",
        key_path=CERTS / "key.pem",
    )
    await server.start()
    print("[harness] Bridge server listening on :3000/:3001")
    print("[harness] Open Excel with the sideloaded add-in now.")

    try:
        await wait_for_addin(server)
        print("[harness] Add-in connected.")

        # Check if spec already exists (resume from Builder phase)
        spec_path = session.run_dir / "model_spec.json"
        if spec_path.exists():
            spec = json.loads(spec_path.read_text())
            await server.send_chat(f"Resuming with existing spec ({len(spec['sheets'])} sheets). Starting Builder...")
            print(f"[harness] Found existing spec. Skipping Planner, jumping to Builder.")
        else:
            brief = await ask_user(server, f"Hi. What would you like to build? Drop input files into {session.input_dir} first, then paste your brief.")
            session.save_brief(brief)
            session.append_chat("user", brief)

            # Preprocess input xlsx files — text dumps only (no screenshots for Planner)
            xlsx_files = preprocess_inputs(session, screenshots=False)
            if xlsx_files:
                await server.send_chat(f"Preprocessed {len(xlsx_files)} input file(s). Starting planning...")
            else:
                await server.send_chat("No input xlsx files found. Starting planning...")

            spec = await run_planner(session, server, brief)
            await server.send_chat(f"Spec complete: {len(spec['sheets'])} sheet(s). Saved to {session.run_dir.name}/model_spec.json")
            print(f"[harness] Spec saved. Planner phase done.")

            # Render screenshots of input files for Builder (visual reference)
            xlsx_files = sorted(session.input_dir.glob("*.xlsx")) + sorted(session.input_dir.glob("*.xls"))
            if xlsx_files:
                await server.send_chat("Rendering input screenshots for Builder...")
                preprocess_inputs(session, screenshots=True)

        await run_builder_loop(session, server, spec)
        print(f"[harness] Builder loop done.")
    finally:
        try:
            await server.send_command("unprotectWorkbook", {})
        except Exception:
            pass
        await server.stop()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[harness] Stopped")
