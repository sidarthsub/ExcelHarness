"""In-process HTTPS + WSS server that lets the harness talk to Excel via the Office.js add-in.

Replaces officejs-prototype/server/relay.py with an embeddable component. The harness
owns this instance so there's no second process and no cross-process IPC.
"""
from __future__ import annotations

import asyncio
import json
import ssl
from pathlib import Path
from typing import Optional

import websockets
from aiohttp import web

ADDIN_DIR = Path(__file__).parent / "officejs-prototype" / "addin"


class BridgeServer:
    def __init__(
        self,
        host: str,
        http_port: int,
        wss_port: int,
        cert_path: Path,
        key_path: Path,
    ):
        self.host = host
        self.http_port = http_port
        self.wss_port = wss_port
        self.cert_path = cert_path
        self.key_path = key_path

        self._runner: Optional[web.AppRunner] = None
        self._wss_server = None
        self._ssl_ctx: Optional[ssl.SSLContext] = None

    async def start(self) -> None:
        self._ssl_ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        self._ssl_ctx.load_cert_chain(self.cert_path, self.key_path)

        app = web.Application()
        app.router.add_get("/{path:.*}", self._handle_static)

        self._runner = web.AppRunner(app)
        await self._runner.setup()
        site = web.TCPSite(self._runner, self.host, self.http_port, ssl_context=self._ssl_ctx)
        await site.start()

        self._wss_server = await websockets.serve(
            self._handle_addin_ws, self.host, self.wss_port, ssl=self._ssl_ctx
        )

    async def stop(self) -> None:
        if self._wss_server is not None:
            self._wss_server.close()
            await self._wss_server.wait_closed()
            self._wss_server = None
        if self._runner is not None:
            await self._runner.cleanup()
            self._runner = None

    async def _handle_static(self, request: web.Request) -> web.StreamResponse:
        path = request.match_info.get("path", "taskpane.html") or "taskpane.html"
        if path == "/":
            path = "taskpane.html"
        file_path = ADDIN_DIR / path
        if file_path.exists() and file_path.is_file():
            resp = web.FileResponse(file_path)
            resp.headers["Cache-Control"] = "no-store"
            return resp
        return web.Response(status=404, text="Not found")

    async def _handle_addin_ws(self, websocket):
        # Minimal stub — real message handling comes in Task 1.4.
        async for _ in websocket:
            pass
