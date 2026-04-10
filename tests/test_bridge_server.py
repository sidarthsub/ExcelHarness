"""Tests for the in-process bridge server that replaces relay.py."""
import asyncio
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
