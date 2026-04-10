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


async def test_inbound_chat_enqueues(server):
    """A {type: 'chat', text: ...} message from the add-in ends up in the ChatQueue."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    import websockets
    async def fake_addin():
        uri = f"wss://{server.host}:{server.wss_port}"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            await ws.send(json.dumps({"type": "chat", "text": "hi from user"}))
            await asyncio.sleep(0.3)  # let server process

    await fake_addin()

    # Drain the queue and confirm the message arrived.
    assert server.chat_queue.drain_all() == ["hi from user"]


async def test_outbound_chat_sends_to_addin(server):
    """server.send_chat(text) pushes a {type: 'chat', ...} message to the connected add-in."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    received = []
    import websockets
    async def fake_addin():
        uri = f"wss://{server.host}:{server.wss_port}"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            msg = await ws.recv()
            received.append(json.loads(msg))

    task = asyncio.create_task(fake_addin())
    await asyncio.sleep(0.2)
    await server.send_chat("hello from harness")
    await asyncio.wait_for(task, timeout=2.0)

    assert len(received) == 1
    assert received[0]["type"] == "chat"
    assert received[0]["text"] == "hello from harness"


async def test_chat_roundtrip_in_process(server):
    """A user chat message goes into the queue, harness code drains it, and a reply lands back in the add-in."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    received_from_harness = []
    import websockets
    async def fake_addin():
        uri = f"wss://{server.host}:{server.wss_port}"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            # Send a user message
            await ws.send(json.dumps({"type": "chat", "text": "ping"}))
            # Wait for the harness's reply
            msg = await ws.recv()
            received_from_harness.append(json.loads(msg))

    task = asyncio.create_task(fake_addin())
    await asyncio.sleep(0.3)

    # Simulate the harness loop: drain, decide, send.
    pending = server.chat_queue.drain_all()
    assert pending == ["ping"]
    await server.send_chat("pong")

    await asyncio.wait_for(task, timeout=2.0)
    assert received_from_harness[0]["type"] == "chat"
    assert received_from_harness[0]["text"] == "pong"


async def test_checkpoint_long_poll(server):
    """Bridge client's POST /api/checkpoint blocks until the harness resolves it."""
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    # Kick off a checkpoint request in a background task
    async def do_checkpoint():
        async with aiohttp.ClientSession() as s:
            async with s.post(
                f"https://{server.host}:{server.http_port}/api/checkpoint",
                json={"description": "Revenue sheet complete"},
                ssl=ssl_ctx,
            ) as resp:
                return await resp.json()

    cp_task = asyncio.create_task(do_checkpoint())
    await asyncio.sleep(0.2)

    # Server should show a pending checkpoint
    pending = server.pop_pending_checkpoint()
    assert pending is not None
    assert pending.description == "Revenue sheet complete"

    # Resolve it
    pending.resolve({"status": "pass"})

    # The HTTP call should now return
    result = await asyncio.wait_for(cp_task, timeout=2.0)
    assert result["status"] == "pass"
