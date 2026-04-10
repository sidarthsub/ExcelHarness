"""Tests for the in-process bridge server that replaces relay.py."""
import asyncio
import json
import ssl
import pytest
import aiohttp
from pathlib import Path

from bridge_server import BridgeServer

CERTS_DIR = Path(__file__).parent.parent / "officejs-prototype" / "certs"


@pytest.fixture
async def server():
    srv = BridgeServer(
        host="localhost",
        http_port=13000,
        wss_port=13001,
        cert_path=CERTS_DIR / "cert.pem",
        key_path=CERTS_DIR / "key.pem",
    )
    await srv.start()
    yield srv
    await srv.stop()


async def test_server_starts_and_stops(server):
    """The server should listen on HTTPS and return taskpane.html on GET /."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE
    async with aiohttp.ClientSession() as session:
        async with session.get("https://localhost:13000/taskpane.html", ssl=ssl_ctx) as resp:
            assert resp.status == 200
            body = await resp.text()
            assert "Excel Harness Bridge" in body


async def test_send_command_to_addin(server):
    """send_command() forwards to the add-in and returns the response."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    fake_addin_task_done = asyncio.Event()

    async def fake_addin():
        import websockets
        uri = f"wss://{server.host}:{server.wss_port}"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            msg = await ws.recv()
            data = json.loads(msg)
            await ws.send(json.dumps({
                "id": data["id"],
                "ok": True,
                "echo": data.get("command"),
            }))
            fake_addin_task_done.set()

    addin_task = asyncio.create_task(fake_addin())
    # Give the add-in a moment to connect.
    await asyncio.sleep(0.2)

    result = await server.send_command("writeValues", {"sheet": "X", "address": "A1", "values": [[1]]})
    assert result["ok"] is True
    assert result["echo"] == "writeValues"

    await asyncio.wait_for(fake_addin_task_done.wait(), timeout=2.0)
    addin_task.cancel()


async def test_http_api_command_endpoint(server):
    """POST /api/command forwards to the add-in and returns the result."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    async def fake_addin():
        import websockets
        uri = f"wss://{server.host}:{server.wss_port}"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            msg = await ws.recv()
            data = json.loads(msg)
            await ws.send(json.dumps({"id": data["id"], "ok": True}))

    addin_task = asyncio.create_task(fake_addin())
    await asyncio.sleep(0.2)

    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"https://{server.host}:{server.http_port}/api/command",
            json={"command": "writeValues", "params": {"sheet": "X", "address": "A1", "values": [[1]]}},
            ssl=ssl_ctx,
        ) as resp:
            assert resp.status == 200
            body = await resp.json()
            assert body["ok"] is True

    addin_task.cancel()
