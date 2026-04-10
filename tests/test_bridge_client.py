"""Tests for bridge.py — the thin HTTP client Builder scripts import."""
import asyncio
import ssl
import json
import pytest
import aiohttp

from bridge_server import BridgeServer
from bridge import Bridge
from pathlib import Path

CERTS_DIR = Path(__file__).parent.parent / "officejs-prototype" / "certs"


@pytest.fixture
async def server_with_fake_addin():
    srv = BridgeServer(
        host="localhost",
        http_port=13002,
        wss_port=13003,
        cert_path=CERTS_DIR / "cert.pem",
        key_path=CERTS_DIR / "key.pem",
    )
    await srv.start()

    # Spawn a fake add-in that echoes every command as {ok: True, command: <name>}.
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    import websockets
    async def fake_addin():
        uri = f"wss://localhost:13003"
        async with websockets.connect(uri, ssl=ssl_ctx) as ws:
            async for msg in ws:
                data = json.loads(msg)
                await ws.send(json.dumps({
                    "id": data["id"],
                    "ok": True,
                    "command": data.get("command"),
                }))

    task = asyncio.create_task(fake_addin())
    await asyncio.sleep(0.2)
    try:
        yield srv
    finally:
        task.cancel()
        await srv.stop()


async def test_bridge_write_values(server_with_fake_addin):
    """Bridge.write_values() sends a writeValues command.

    Note: Bridge is a synchronous client (Builder scripts are sync), but the
    test is async and wraps the call in asyncio.to_thread so the fixture's
    in-process aiohttp server can service the request on the event loop while
    the blocking requests.post runs in a worker thread.
    """
    b = Bridge(base_url="https://localhost:13002", verify_tls=False)
    result = await asyncio.to_thread(b.write_values, "Sheet1", "A1:B1", [[1, 2]])
    assert result["ok"] is True
    assert result["command"] == "writeValues"


async def test_bridge_write_formulas(server_with_fake_addin):
    b = Bridge(base_url="https://localhost:13002", verify_tls=False)
    result = await asyncio.to_thread(b.write_formulas, "Sheet1", "A1", [["=1+1"]])
    assert result["ok"] is True
    assert result["command"] == "writeFormulas"
