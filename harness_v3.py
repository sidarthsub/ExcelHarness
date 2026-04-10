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
) -> str:
    """Run a single-turn Claude Agent SDK query and return the text of the final assistant response.

    For multi-turn Planner flows (ambiguity Q&A -> spec), call this multiple times
    with accumulated user_messages.
    """
    options = ClaudeAgentOptions(
        system_prompt=system_prompt,
        allowed_tools=allowed_tools,
        cwd=str(cwd),
    )

    # For now we concatenate all user turns into one prompt. The SDK's query() API
    # is single-turn — for real multi-turn we need a ClaudeSDKClient (v2 harness uses
    # it). We'll upgrade to a session-based client if single-turn proves insufficient.
    combined = "\n\n---\n\n".join(user_messages)

    final_text = ""
    async for message in query(prompt=combined, options=options):
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    final_text = block.text
    return final_text


async def wait_for_addin(server: BridgeServer, timeout: float = 60.0) -> None:
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
            return pending[0]
        await asyncio.sleep(0.3)


async def run_planner(session: Session, server: BridgeServer, brief: str) -> dict:
    """Three-pass Planner: ambiguities -> spec -> self-review."""
    prompt = (AGENTS_DIR / "planner_v3.md").read_text()
    schema = json.loads((ROOT / "schemas" / "model_spec.schema.json").read_text())

    # --- Pass 1: Ambiguities ---
    pass1_user = (
        f"BRIEF:\n{brief}\n\n"
        "Run Pass 1 (ambiguity detection). Output a JSON array of questions, "
        "or an empty array if the brief is fully specified. Respond with ONLY the JSON array."
    )

    pass1_text = await run_agent_session(
        system_prompt=prompt,
        user_messages=[pass1_user],
        allowed_tools=["Read", "Glob", "Grep"],
        cwd=ROOT,
    )

    try:
        questions = json.loads(pass1_text.strip())
    except json.JSONDecodeError:
        # Try to extract JSON from markdown code block
        import re
        m = re.search(r"\[.*?\]", pass1_text, re.DOTALL)
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
    clarif_text = "\n".join(f"- {k}: {v}" for k, v in clarifications.items())
    pass2_user = (
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
    )

    # --- Load, validate, optionally retry on schema failure ---
    spec_path = session.run_dir / "model_spec.json"
    if not spec_path.exists():
        raise RuntimeError("Planner did not write model_spec.json")

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
        jsonschema.validate(instance=spec, schema=schema)

    # --- Pass 3: Self-review ---
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
    )
    # Reload and revalidate in case pass 3 rewrote.
    spec = json.loads(spec_path.read_text())
    jsonschema.validate(instance=spec, schema=schema)

    return spec


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

        brief = await ask_user(server, "Hi. What would you like to build? Paste your brief (and attach files separately to runs/<session>/input/).")
        session.save_brief(brief)
        session.append_chat("user", brief)

        await server.send_chat("Thanks. Starting planning...")
        spec = await run_planner(session, server, brief)
        await server.send_chat(f"Spec complete: {len(spec['sheets'])} sheet(s). Saved to {session.run_dir.name}/model_spec.json")
        print(f"[harness] Spec saved. Planner phase done.")
    finally:
        await server.stop()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[harness] Stopped")
