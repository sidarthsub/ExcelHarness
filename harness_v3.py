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

ROOT = Path(__file__).parent
CERTS = ROOT / "officejs-prototype" / "certs"
AGENTS_DIR = ROOT / "agents"


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


async def run_planner(session: Session, brief: str) -> dict:
    """Run the planning phase and return the validated spec.

    For now this is a stub that returns a hardcoded spec. Task 3.6 wires in the
    real Claude Agent SDK invocation.
    """
    spec = {
        "intent": f"Placeholder spec built from brief: {brief[:80]}",
        "sheets": [
            {"name": "Inputs", "purpose": "Assumptions.", "must_contain": [], "data_sources": []},
        ],
        "constraints": [],
        "out_of_scope": [],
    }
    session.save_spec(spec)
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
        spec = await run_planner(session, brief)
        await server.send_chat(f"Spec complete: {len(spec['sheets'])} sheet(s). Saved to {session.run_dir.name}/model_spec.json")
        print(f"[harness] Spec saved. Planner phase done.")
    finally:
        await server.stop()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[harness] Stopped")
