# Live Excel Architecture Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Rebuild the harness around an always-live Excel session via Office.js, replacing the openpyxl + dump pipeline end-to-end.

**Architecture:** Two components — an Office.js add-in in Excel's taskpane, and a single Python harness process that serves the bridge, runs the agents, and owns the chat channel. See [2026-04-10-live-excel-architecture.md](2026-04-10-live-excel-architecture.md) for the full design.

**Tech Stack:** Python 3.13 (aiohttp, websockets, claude-agent-sdk, jsonschema), Office.js (Excel APIs), LibreOffice (snapshot rendering only), git (version control at checkpoints).

**Branch:** `v3-live-excel` off `main`

**Working directory:** Assume all relative paths are from `/Users/sidsub/Documents/ExcelHarness` unless specified.

---

## How to execute this plan

Each task follows TDD where feasible: write a failing test, run it, implement, run again, commit. Some tasks (UI work in the taskpane, prompt iteration) don't fit TDD cleanly — those have explicit manual verification steps instead. Commit after every task. If a task says "verify by running X and confirming output matches Y", do not skip that verification — the plan assumes each step is proven before the next begins.

When a task says "write the agent prompt", the prompt content is a starting point. Iterate on it if the first run produces bad output, but commit the starting point first so the iteration history is visible in git.

---

## Phase 0: Branch setup and prototype commit

### Task 0.1: Create the v3 branch

**Files:** none modified; branch-level change only.

**Step 1:** Verify you're on a clean tree for the tracked files you care about.

Run: `git status --short`

Expected: The `officejs-prototype/` directory shows as untracked, `harness.py` and `.claude/settings.json` show as modified, and `architecture.html`, `evals/`, `status.json` are untracked. These will stay as-is — we're only going to commit the prototype directory in this phase.

**Step 2:** Create the branch off `main`.

Run: `git checkout -b v3-live-excel main`

Expected: `Switched to a new branch 'v3-live-excel'`

Note: You'll see the same uncommitted changes on the new branch because they're in the working tree. That's fine.

**Step 3:** Commit nothing yet. The branch exists now; we'll commit in 0.2.

### Task 0.2: Commit the prototype as the starting point

**Files:**
- Add: `officejs-prototype/` (entire directory currently untracked)

**Step 1:** Stage the prototype directory.

Run: `git add officejs-prototype/`

**Step 2:** Verify what got staged.

Run: `git status --short`

Expected: A block of `A` (added) entries under `officejs-prototype/`: the `addin/` files (`taskpane.html`, `taskpane_v8.html`, `manifest.xml`, three PNG icons), `server/relay.py`, `certs/cert.pem`, and `certs/key.pem`.

**Step 3:** Verify the certs we're about to commit aren't private long-lived identities. These are Microsoft's `office-addin-dev-certs` localhost-only dev certs, valid only for `localhost`. They're safe to commit for a prototype but add a note to the commit message.

**Step 4:** Commit.

Run:
```bash
git commit -m "$(cat <<'EOF'
Commit Office.js bridge prototype as v3 starting point

Prototype of the add-in + WebSocket relay that proved the feasibility
of live Excel control. Moves officejs-prototype/ from untracked into
git so it can serve as the foundation for the v3 rewrite.

Certs in officejs-prototype/certs/ are localhost-only dev certs from
Microsoft's office-addin-dev-certs tooling. Safe to commit for local
development; do not reuse for anything public-facing.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>
EOF
)"
```

Expected: `[v3-live-excel ...] Commit Office.js bridge prototype as v3 starting point`

### Task 0.3: Delete the deprecated `taskpane_v8.html`

The prototype carries cache-busting duplicates (`taskpane.html` + `taskpane_v8.html`). We keep `taskpane.html` as the canonical file going forward. The manifest currently points at `taskpane_v8.html`; fix that now.

**Files:**
- Delete: `officejs-prototype/addin/taskpane_v8.html`
- Modify: `officejs-prototype/addin/manifest.xml`

**Step 1:** Verify the two files are identical so we're not deleting newer work.

Run: `diff officejs-prototype/addin/taskpane.html officejs-prototype/addin/taskpane_v8.html`

Expected: no output (files are identical). If they differ, stop and reconcile before deleting.

**Step 2:** Delete the duplicate.

Run: `git rm officejs-prototype/addin/taskpane_v8.html`

**Step 3:** Update the manifest to point at `taskpane.html`.

Edit `officejs-prototype/addin/manifest.xml`, find the two references to `taskpane_v8.html` (the `<SourceLocation>` under `<DefaultSettings>` and the `<bt:Url id="Taskpane.Url">` inside `<Resources>`) and change both to `taskpane.html`.

Expected edit result: two occurrences of `taskpane_v8.html` become `taskpane.html`.

**Step 4:** Verify.

Run: `grep -c "taskpane" officejs-prototype/addin/manifest.xml`

Expected: all references point at `taskpane.html`, zero remaining `v8` references.

Run: `grep "v8" officejs-prototype/addin/manifest.xml`

Expected: no output.

**Step 5:** Commit.

```bash
git add officejs-prototype/addin/manifest.xml officejs-prototype/addin/taskpane_v8.html
git commit -m "$(cat <<'EOF'
Collapse cache-busting taskpane duplicate

Remove taskpane_v8.html (identical copy of taskpane.html kept around
to dodge Excel's WebView cache during prototype iteration). Point
manifest.xml back at taskpane.html. Cache-busting is no longer needed
now that we have a stable version.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 0.4: Add a `.gitignore` entry for per-session runs

**Files:**
- Modify: `.gitignore` (create if missing)

**Step 1:** Check existing gitignore.

Run: `cat .gitignore 2>/dev/null || echo "no gitignore yet"`

**Step 2:** Add entries for v3-specific transient artifacts. Create or append to `.gitignore`:

```
# v3 per-session artifacts
runs/
models/*.xlsx
!models/.gitkeep

# Python
__pycache__/
*.pyc
.venv/

# Office sideload temp files
/tmp/Excel add-in *.xlsx
```

**Step 3:** Create an empty `models/.gitkeep` so the directory exists in git.

Run: `mkdir -p models && touch models/.gitkeep`

**Step 4:** Commit.

```bash
git add .gitignore models/.gitkeep
git commit -m "Ignore per-session runs and model xlsx outputs

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

---

## Phase 1: Single-process harness owns the bridge

Goal: Replace the standalone `relay.py` with an in-process bridge owned by `harness.py`. Write `bridge.py` (the HTTP client library for Builder scripts). Prove end-to-end that a hand-written Python script using `bridge.py` can build the prototype test sheet via live Excel.

**Important architectural note:** The current `harness.py` (925 lines) is the v2 implementation. We are NOT modifying it. We're writing a new `harness.py` from scratch in small pieces. To avoid destroying v2 during incremental work, we'll write the new code in a new file `harness_v3.py` and rename it to `harness.py` at the end of Phase 6 when v2 is deleted.

### Task 1.1: Set up the test directory and pytest config

**Files:**
- Create: `tests/__init__.py` (empty)
- Create: `tests/conftest.py`
- Create: `pytest.ini`

**Step 1:** Check if pytest is already installed.

Run: `python3 -c "import pytest; print(pytest.__version__)" 2>&1`

Expected: a version number. If not installed, run `pip3 install pytest pytest-asyncio aiohttp websockets jsonschema`.

**Step 2:** Create the test package.

```bash
mkdir -p tests
touch tests/__init__.py
```

**Step 3:** Create `pytest.ini`:

```ini
[pytest]
testpaths = tests
asyncio_mode = auto
python_files = test_*.py
python_classes = Test*
python_functions = test_*
```

**Step 4:** Create `tests/conftest.py` with a basic fixture scaffold:

```python
"""Shared fixtures for the v3 harness test suite."""
import pytest
from pathlib import Path

HARNESS_ROOT = Path(__file__).parent.parent


@pytest.fixture
def harness_root():
    return HARNESS_ROOT
```

**Step 5:** Verify pytest discovers zero tests (no tests yet, just config).

Run: `python3 -m pytest --collect-only 2>&1`

Expected: `collected 0 items`

**Step 6:** Commit.

```bash
git add tests/ pytest.ini
git commit -m "Scaffold pytest for v3 harness

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.2: Write the test for `BridgeServer.start()` / `stop()` lifecycle

The core of Phase 1 is a `BridgeServer` class that owns the HTTP+WSS listeners. We'll test its lifecycle first, then build the command dispatch on top.

**Files:**
- Create: `tests/test_bridge_server.py`

**Step 1:** Write the failing test. `tests/test_bridge_server.py`:

```python
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
    """The server should listen on both configured ports after start()."""
    # HTTPS health check: GET / returns 200 with the add-in static files.
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE
    async with aiohttp.ClientSession() as session:
        async with session.get("https://localhost:13000/taskpane.html", ssl=ssl_ctx) as resp:
            assert resp.status == 200
            body = await resp.text()
            assert "Excel Harness Bridge" in body
```

**Step 2:** Run the test to verify it fails.

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: `ImportError: No module named 'bridge_server'` or `ModuleNotFoundError`. The import itself fails, so the test fails fast.

### Task 1.3: Implement `BridgeServer` start/stop

**Files:**
- Create: `bridge_server.py`

**Step 1:** Write the minimal implementation to make the lifecycle test pass. `bridge_server.py`:

```python
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
```

**Step 2:** Run the test.

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: PASS. If the test fails because of cert loading, confirm `openssl x509 -in officejs-prototype/certs/cert.pem -noout -dates` shows the cert is valid. If there's a port conflict (13000 already in use), change the ports in the fixture and `BridgeServer` test to something free.

**Step 3:** Commit.

```bash
git add bridge_server.py tests/test_bridge_server.py
git commit -m "Add BridgeServer skeleton with HTTPS static file serving

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.4: Write the test for command dispatch via in-process API

The bridge server needs to accept commands from in-process Python code (not HTTP — the harness calls directly) and forward them to the add-in over WSS. We'll test this with a fake add-in client that echoes back.

**Files:**
- Modify: `tests/test_bridge_server.py`

**Step 1:** Add a new test. Append to `tests/test_bridge_server.py`:

```python
async def test_send_command_to_addin(server):
    """send_command() forwards to the add-in and returns the response."""
    # Fake add-in: connect via WSS, echo every command back with the same id.
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
```

Also add `import json` at the top if not already present.

**Step 2:** Run the test.

Run: `python3 -m pytest tests/test_bridge_server.py::test_send_command_to_addin -v 2>&1`

Expected: FAIL with `AttributeError: 'BridgeServer' object has no attribute 'send_command'`.

### Task 1.5: Implement `send_command()` with an add-in connection registry

**Files:**
- Modify: `bridge_server.py`

**Step 1:** Add command dispatch to `BridgeServer`. Replace the `_handle_addin_ws` stub and add `send_command`:

```python
# Add at the top of the class __init__:
#     self._addin_ws = None
#     self._pending: dict[str, asyncio.Future] = {}
#     self._msg_counter = 0

async def _handle_addin_ws(self, websocket):
    self._addin_ws = websocket
    try:
        async for message in websocket:
            data = json.loads(message)
            msg_id = data.get("id")
            if msg_id and msg_id in self._pending:
                self._pending[msg_id].set_result(data)
    except websockets.exceptions.ConnectionClosed:
        pass
    finally:
        if self._addin_ws is websocket:
            self._addin_ws = None

async def send_command(self, command: str, params: dict | None = None, timeout: float = 30.0) -> dict:
    if self._addin_ws is None:
        return {"error": "no add-in connected"}
    self._msg_counter += 1
    msg_id = f"cmd-{self._msg_counter}"
    payload = {"id": msg_id, "command": command, "params": params or {}}

    future = asyncio.get_event_loop().create_future()
    self._pending[msg_id] = future
    try:
        await self._addin_ws.send(json.dumps(payload))
        return await asyncio.wait_for(future, timeout=timeout)
    except asyncio.TimeoutError:
        return {"error": "timeout"}
    finally:
        self._pending.pop(msg_id, None)
```

Don't forget to update `__init__` to initialize `self._addin_ws`, `self._pending`, and `self._msg_counter`. Add `import json` at the top of the file.

**Step 2:** Run both tests.

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: both PASS.

**Step 3:** Commit.

```bash
git add bridge_server.py tests/test_bridge_server.py
git commit -m "Add command dispatch to BridgeServer

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.6: Write the test for HTTP command endpoint

Builder scripts will POST commands to `/api/command` on the harness's own HTTP server. Write that test.

**Files:**
- Modify: `tests/test_bridge_server.py`

**Step 1:** Add the test:

```python
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
```

**Step 2:** Run:

Run: `python3 -m pytest tests/test_bridge_server.py::test_http_api_command_endpoint -v 2>&1`

Expected: FAIL — route doesn't exist yet.

### Task 1.7: Implement the HTTP command endpoint

**Files:**
- Modify: `bridge_server.py`

**Step 1:** Register the `/api/command` POST route before the catch-all static route. In `start()`:

```python
app.router.add_post("/api/command", self._handle_api_command)
app.router.add_get("/{path:.*}", self._handle_static)
```

Add the handler:

```python
async def _handle_api_command(self, request: web.Request) -> web.Response:
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON"}, status=400)
    command = body.get("command")
    params = body.get("params") or {}
    if not command:
        return web.json_response({"error": "missing command"}, status=400)
    result = await self.send_command(command, params)
    return web.json_response(result)
```

**Step 2:** Run tests.

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: all three tests PASS.

**Step 3:** Commit.

```bash
git add bridge_server.py tests/test_bridge_server.py
git commit -m "Expose /api/command HTTP endpoint for in-process clients

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.8: Write `bridge.py` — the thin Python client

**Files:**
- Create: `tests/test_bridge_client.py`
- Create: `bridge.py`

**Step 1:** Write the failing test. `tests/test_bridge_client.py`:

```python
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


def test_bridge_write_values(server_with_fake_addin):
    """Bridge.write_values() sends a writeValues command."""
    b = Bridge(base_url="https://localhost:13002", verify_tls=False)
    result = b.write_values("Sheet1", "A1:B1", [[1, 2]])
    assert result["ok"] is True
    assert result["command"] == "writeValues"


def test_bridge_write_formulas(server_with_fake_addin):
    b = Bridge(base_url="https://localhost:13002", verify_tls=False)
    result = b.write_formulas("Sheet1", "A1", [["=1+1"]])
    assert result["ok"] is True
    assert result["command"] == "writeFormulas"
```

Note: these tests use sync Bridge calls inside async fixtures. That's intentional — Builder scripts are synchronous, not async. The Bridge client runs its own event loop or uses `requests` instead of `aiohttp`. Use `requests` to keep it simple.

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_bridge_client.py -v 2>&1`

Expected: FAIL with `ModuleNotFoundError: No module named 'bridge'`.

### Task 1.9: Implement `bridge.py`

**Files:**
- Create: `bridge.py`

**Step 1:** Write the client. `bridge.py`:

```python
"""Synchronous HTTP client that Builder scripts import to drive live Excel.

This is the interface the Builder agent writes code against. Keep it thin —
one function per bridge command, no convention encoding, no helpers.
"""
from __future__ import annotations

import requests
import urllib3


class BridgeError(Exception):
    pass


class Bridge:
    def __init__(self, base_url: str = "https://localhost:3000", verify_tls: bool = True, timeout: float = 30.0):
        self.base_url = base_url.rstrip("/")
        self.verify = verify_tls
        self.timeout = timeout
        if not verify_tls:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def _call(self, command: str, params: dict | None = None) -> dict:
        resp = requests.post(
            f"{self.base_url}/api/command",
            json={"command": command, "params": params or {}},
            verify=self.verify,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        body = resp.json()
        if "error" in body and not body.get("ok"):
            raise BridgeError(f"{command} failed: {body['error']}")
        return body

    # --- Sheet management ---
    def create_sheet(self, name: str) -> dict:
        return self._call("createSheet", {"name": name})

    def delete_sheet(self, name: str) -> dict:
        return self._call("deleteSheet", {"name": name})

    def get_sheet_names(self) -> dict:
        return self._call("getSheetNames", {})

    # --- Values and formulas ---
    def write_values(self, sheet: str, address: str, values: list) -> dict:
        return self._call("writeValues", {"sheet": sheet, "address": address, "values": values})

    def write_formulas(self, sheet: str, address: str, formulas: list) -> dict:
        return self._call("writeFormulas", {"sheet": sheet, "address": address, "formulas": formulas})

    def read_values(self, sheet: str, address: str) -> dict:
        return self._call("readValues", {"sheet": sheet, "address": address})

    def read_format(self, sheet: str, address: str) -> dict:
        return self._call("readFormat", {"sheet": sheet, "address": address})

    # --- Formatting ---
    def format_range(self, sheet: str, address: str, format: dict) -> dict:
        return self._call("formatRange", {"sheet": sheet, "address": address, "format": format})

    def set_number_format(self, sheet: str, address: str, format: str) -> dict:
        return self._call("setNumberFormat", {"sheet": sheet, "address": address, "format": format})

    def set_column_widths(self, sheet: str, columns: dict) -> dict:
        return self._call("setColumnWidths", {"sheet": sheet, "columns": columns})

    def set_row_heights(self, sheet: str, rows: dict) -> dict:
        return self._call("setRowHeights", {"sheet": sheet, "rows": rows})

    def merge_cells(self, sheet: str, address: str) -> dict:
        return self._call("mergeCells", {"sheet": sheet, "address": address})

    def freeze_rows(self, sheet: str, count: int) -> dict:
        return self._call("freezeRows", {"sheet": sheet, "count": count})

    def clear_range(self, sheet: str, address: str) -> dict:
        return self._call("clearRange", {"sheet": sheet, "address": address})

    # --- Charts ---
    def create_chart(self, sheet: str, **kwargs) -> dict:
        return self._call("createChart", {"sheet": sheet, **kwargs})

    # --- Workbook settings ---
    def set_iterative_calculation(self, enabled: bool, max_iteration: int = 100, max_change: float = 0.001) -> dict:
        return self._call("setIterativeCalculation", {
            "enabled": enabled,
            "maxIteration": max_iteration,
            "maxChange": max_change,
        })

    def create_table(self, sheet: str, header_address: str, name: str, rows: list, style: str | None = None) -> dict:
        params = {"sheet": sheet, "headerAddress": header_address, "name": name, "rows": rows}
        if style:
            params["style"] = style
        return self._call("createTable", params)

    # --- Batch ---
    def batch(self, commands: list) -> dict:
        return self._call("batch", {"commands": commands})
```

**Step 2:** Run tests.

Run: `python3 -m pytest tests/test_bridge_client.py -v 2>&1`

Expected: both tests PASS.

**Step 3:** Commit.

```bash
git add bridge.py tests/test_bridge_client.py
git commit -m "Add bridge.py — thin HTTP client for Builder scripts

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.10: Add new bridge commands to the add-in (dumpSheet, protect/unprotect, saveSnapshot, setStatus)

The prototype taskpane is missing four commands the new architecture needs. Add them now so Phase 1's end-to-end test can exercise them.

**Files:**
- Modify: `officejs-prototype/addin/taskpane.html`

**Step 1:** Locate the `handlers` object in the JavaScript section of `taskpane.html`. It's currently around lines 45-300.

**Step 2:** Add these four handlers. Place them after the existing `batch` handler:

```javascript
async dumpSheet({ sheet }) {
    return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(sheet);
        const used = ws.getUsedRange(true /* onlyValues=false, include all */);
        used.load(["address", "values", "formulas", "numberFormat", "rowCount", "columnCount"]);
        await ctx.sync();
        return {
            ok: true,
            sheet: sheet,
            address: used.address,
            dimensions: { rows: used.rowCount, cols: used.columnCount },
            values: used.values,
            formulas: used.formulas,
            numberFormat: used.numberFormat,
        };
    });
},

async protectWorkbook({ password }) {
    return Excel.run(async (ctx) => {
        // Protect every worksheet so users can't type into cells.
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();
        for (const ws of sheets.items) {
            ws.protection.protect({}, password);
        }
        await ctx.sync();
        return { ok: true };
    });
},

async unprotectWorkbook({ password }) {
    return Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        await ctx.sync();
        for (const ws of sheets.items) {
            ws.protection.unprotect(password);
        }
        await ctx.sync();
        return { ok: true };
    });
},

async saveSnapshot({ path }) {
    // Workbook.save() writes to the default location; we return the file contents
    // as base64 so the harness can write it wherever it wants.
    return Excel.run(async (ctx) => {
        const file = ctx.workbook.getFileInBase64();
        await ctx.sync();
        return { ok: true, base64: file.value };
    });
},

async setStatus({ text }) {
    // Updates the status line in the taskpane. No Excel API call — just DOM.
    const el = document.getElementById("status");
    if (el) el.textContent = text;
    return { ok: true };
},
```

**Step 3:** Add the corresponding methods to `bridge.py`:

```python
def dump_sheet(self, sheet: str) -> dict:
    return self._call("dumpSheet", {"sheet": sheet})

def protect_workbook(self, password: str | None = None) -> dict:
    return self._call("protectWorkbook", {"password": password})

def unprotect_workbook(self, password: str | None = None) -> dict:
    return self._call("unprotectWorkbook", {"password": password})

def save_snapshot(self) -> dict:
    return self._call("saveSnapshot", {})

def set_status(self, text: str) -> dict:
    return self._call("setStatus", {"text": text})
```

**Step 4:** Verify the add-in JavaScript is syntactically valid by opening the file in a browser or running it through a JS linter. Quick sanity check:

Run: `python3 -c "import re; open('officejs-prototype/addin/taskpane.html').read()" 2>&1`

Expected: no error. (This just confirms the file is readable; there's no cheap JS syntax check from Python.)

**Step 5:** Commit.

```bash
git add officejs-prototype/addin/taskpane.html bridge.py
git commit -m "Add dumpSheet, protect/unprotect, saveSnapshot, setStatus commands

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.11: End-to-end smoke test with real Excel

This is a **manual verification** task. The test suite doesn't cover live Excel; you'll run it by hand to prove Phase 1 works.

**Files:** none modified.

**Step 1:** Write a small smoke test script. Create `scripts/phase1_smoke.py`:

```python
"""Phase 1 smoke test: hand-written script that exercises the bridge against live Excel.

Before running:
  1. Start the bridge server (Task 1.12 shows how).
  2. Open Excel with the sideloaded add-in.
  3. Wait for the taskpane to show "Connected".

Then run: python3 scripts/phase1_smoke.py
"""
from bridge import Bridge

b = Bridge(base_url="https://localhost:3000", verify_tls=False)

# Basic round-trip
print("Creating sheet...")
b.create_sheet("Phase1 Smoke")

print("Writing values...")
b.write_values("Phase1 Smoke", "A1:B3", [["name", "value"], ["a", 1], ["b", 2]])

print("Writing formula...")
b.write_formulas("Phase1 Smoke", "B4", [["=SUM(B2:B3)"]])

print("Reading back...")
result = b.read_values("Phase1 Smoke", "B4")
print(f"  B4 = {result}")

print("Dumping sheet...")
dump = b.dump_sheet("Phase1 Smoke")
print(f"  dimensions: {dump.get('dimensions')}")
print(f"  values: {dump.get('values')}")

print("Done.")
```

**Step 2:** Start the bridge server. There's no CLI wrapper yet, so use a Python one-liner:

```bash
python3 -c "
import asyncio
from pathlib import Path
from bridge_server import BridgeServer
async def main():
    srv = BridgeServer('localhost', 3000, 3001, Path('officejs-prototype/certs/cert.pem'), Path('officejs-prototype/certs/key.pem'))
    await srv.start()
    print('[bridge] HTTPS on https://localhost:3000, WSS on wss://localhost:3001')
    await asyncio.Future()
asyncio.run(main())
" &
```

Save the background PID for later cleanup.

**Step 3:** Sideload the add-in into Excel:

Run: `npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml`

Expected: Excel opens with the add-in loaded. Taskpane shows "Connected".

**Step 4:** Run the smoke test:

Run: `python3 scripts/phase1_smoke.py`

Expected output:
```
Creating sheet...
Writing values...
Writing formula...
Reading back...
  B4 = {'ok': True, 'values': [[3]]}
Dumping sheet...
  dimensions: {'rows': 4, 'cols': 2}
  values: [['name', 'value'], ['a', 1], ['b', 2], ['', 3]]
Done.
```

Visually verify in Excel that the "Phase1 Smoke" sheet exists with the expected data and `B4` shows `3` (the SUM result).

**Step 5:** Stop the background bridge server:

Run: `kill $(lsof -ti :3000) 2>&1`

**Step 6:** Commit the smoke test script.

```bash
git add scripts/phase1_smoke.py
git commit -m "Add Phase 1 smoke test script for manual live-Excel verification

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 1.12: Add a CLI entry point for the bridge server

**Files:**
- Create: `run_bridge.py`

**Step 1:** Write a tiny standalone runner. `run_bridge.py`:

```python
#!/usr/bin/env python3
"""Start the bridge server as a standalone process (useful during Phase 1-2 dev).

Later phases integrate this into harness_v3.py and you won't run it separately.
"""
import asyncio
from pathlib import Path

from bridge_server import BridgeServer

ROOT = Path(__file__).parent
CERTS = ROOT / "officejs-prototype" / "certs"


async def main():
    srv = BridgeServer(
        host="localhost",
        http_port=3000,
        wss_port=3001,
        cert_path=CERTS / "cert.pem",
        key_path=CERTS / "key.pem",
    )
    await srv.start()
    print("[bridge] HTTPS on https://localhost:3000")
    print("[bridge] WSS on wss://localhost:3001")
    print("[bridge] Press Ctrl+C to stop")
    try:
        await asyncio.Future()
    except asyncio.CancelledError:
        pass
    finally:
        await srv.stop()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n[bridge] Stopped")
```

**Step 2:** Make it executable.

Run: `chmod +x run_bridge.py`

**Step 3:** Smoke-test it one more time with the CLI.

```bash
python3 run_bridge.py &
sleep 2
curl -sk https://localhost:3000/taskpane.html | head -5
kill $(lsof -ti :3000)
```

Expected: The `curl` returns the first 5 lines of taskpane.html. If it does, the CLI works.

**Step 4:** Commit.

```bash
git add run_bridge.py
git commit -m "Add standalone bridge server CLI entry point

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 1 done.** Bridge server in-process, `bridge.py` client working, new commands added, smoke test verified against live Excel.

---

## Phase 2: Chat channel

Goal: Add the chat UI to the taskpane, extend the WSS protocol with chat message types, and give the bridge server an in-memory chat FIFO. No agents yet — this is pure plumbing.

### Task 2.1: Write the test for `ChatQueue`

**Files:**
- Create: `tests/test_chat_queue.py`

**Step 1:** Create the failing test. `tests/test_chat_queue.py`:

```python
"""Tests for the in-memory chat FIFO used by the harness."""
import pytest

from chat_queue import ChatQueue


def test_enqueue_and_drain():
    q = ChatQueue()
    q.enqueue("hello")
    q.enqueue("world")
    drained = q.drain_all()
    assert drained == ["hello", "world"]


def test_drain_when_empty():
    q = ChatQueue()
    assert q.drain_all() == []


def test_drain_clears_queue():
    q = ChatQueue()
    q.enqueue("a")
    q.drain_all()
    assert q.drain_all() == []


def test_thread_safe_concurrent_writes():
    """enqueue and drain should not lose messages under concurrent writes."""
    import threading
    q = ChatQueue()
    N = 1000

    def writer():
        for i in range(N):
            q.enqueue(f"msg-{i}")

    threads = [threading.Thread(target=writer) for _ in range(4)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    all_msgs = q.drain_all()
    assert len(all_msgs) == N * 4
```

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_chat_queue.py -v 2>&1`

Expected: FAIL with `ModuleNotFoundError: No module named 'chat_queue'`.

### Task 2.2: Implement `ChatQueue`

**Files:**
- Create: `chat_queue.py`

**Step 1:** Write minimal implementation. `chat_queue.py`:

```python
"""Thread-safe FIFO for inbound chat messages.

The add-in sends chat messages over WSS; the WSS handler calls enqueue().
The harness's Builder loop calls drain_all() between turns.
"""
import threading


class ChatQueue:
    def __init__(self):
        self._messages: list[str] = []
        self._lock = threading.Lock()

    def enqueue(self, text: str) -> None:
        with self._lock:
            self._messages.append(text)

    def drain_all(self) -> list[str]:
        with self._lock:
            drained = self._messages[:]
            self._messages.clear()
            return drained
```

**Step 2:** Run tests:

Run: `python3 -m pytest tests/test_chat_queue.py -v 2>&1`

Expected: all PASS.

**Step 3:** Commit.

```bash
git add chat_queue.py tests/test_chat_queue.py
git commit -m "Add ChatQueue: thread-safe in-memory FIFO for inbound chat

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 2.3: Extend the WSS protocol with chat message types

**Files:**
- Modify: `bridge_server.py`
- Modify: `tests/test_bridge_server.py`

**Step 1:** Write the failing test. Append to `tests/test_bridge_server.py`:

```python
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
```

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1 | tail -30`

Expected: new tests FAIL (`chat_queue` attribute missing, `send_chat` method missing).

**Step 3:** Implement it. Modify `BridgeServer.__init__` to create a `ChatQueue`:

```python
from chat_queue import ChatQueue
# ...
def __init__(self, ...):
    # existing fields ...
    self.chat_queue = ChatQueue()
```

Modify `_handle_addin_ws` to dispatch by message type:

```python
async def _handle_addin_ws(self, websocket):
    self._addin_ws = websocket
    try:
        async for message in websocket:
            data = json.loads(message)
            msg_type = data.get("type")
            if msg_type == "chat":
                self.chat_queue.enqueue(data.get("text", ""))
                continue
            # Command response
            msg_id = data.get("id")
            if msg_id and msg_id in self._pending:
                self._pending[msg_id].set_result(data)
    except websockets.exceptions.ConnectionClosed:
        pass
    finally:
        if self._addin_ws is websocket:
            self._addin_ws = None

async def send_chat(self, text: str) -> None:
    if self._addin_ws is None:
        return
    await self._addin_ws.send(json.dumps({"type": "chat", "text": text}))
```

**Step 4:** Run tests:

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: all tests PASS (5 total at this point).

**Step 5:** Commit.

```bash
git add bridge_server.py tests/test_bridge_server.py
git commit -m "Add chat message types to WSS protocol

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 2.4: Add chat UI to the taskpane

**Files:**
- Modify: `officejs-prototype/addin/taskpane.html`

**Step 1:** This is a UI change, not TDD. Read the current file:

Run: `wc -l officejs-prototype/addin/taskpane.html`

**Step 2:** Replace the current body HTML with a chat-first layout. Find the `<body>` tag and replace the content between `<body>` and `<script>` with:

```html
<body>
    <div id="status">Initializing...</div>
    <button id="runTest" onclick="runTest()" disabled>Run Test</button>
    <div id="chatLog"></div>
    <div id="chatInputRow">
        <input type="text" id="chatInput" placeholder="Type a message..." disabled />
        <button id="chatSend" onclick="sendChat()" disabled>Send</button>
    </div>
    <details id="debugLog">
        <summary>Debug log</summary>
        <div id="log"></div>
    </details>
```

**Step 3:** Update the `<style>` block to add chat styling. Inside the `<style>` block, add:

```css
#chatLog { background: #0f0f0f; padding: 10px; border-radius: 4px; min-height: 200px; max-height: 400px; overflow-y: auto; margin: 10px 0; font-family: -apple-system, sans-serif; }
.chat-msg { padding: 6px 10px; margin: 4px 0; border-radius: 6px; max-width: 90%; }
.chat-user { background: #1f3864; color: #fff; margin-left: auto; text-align: right; }
.chat-agent { background: #2a2a2a; color: #ccc; }
.chat-system { background: #3a2a1a; color: #e0a050; font-style: italic; }
#chatInputRow { display: flex; gap: 8px; }
#chatInput { flex: 1; background: #1a1a1a; color: #ccc; border: 1px solid #333; padding: 8px; border-radius: 4px; }
#chatInput:disabled { opacity: 0.5; }
#chatSend { background: #2F5496; color: #fff; border: none; padding: 8px 16px; border-radius: 4px; cursor: pointer; }
#chatSend:disabled { background: #555; cursor: not-allowed; }
details#debugLog summary { cursor: pointer; color: #888; margin-top: 12px; }
```

**Step 4:** Add chat JavaScript. Inside the `<script>` block, near the top (after `log()` and `setStatus()`), add:

```javascript
function addChatMessage(text, role) {
    const log = document.getElementById("chatLog");
    const msg = document.createElement("div");
    msg.className = `chat-msg chat-${role}`;
    msg.textContent = text;
    log.appendChild(msg);
    log.scrollTop = log.scrollHeight;
}

function sendChat() {
    const input = document.getElementById("chatInput");
    const text = input.value.trim();
    if (!text || !ws || ws.readyState !== WebSocket.OPEN) return;
    ws.send(JSON.stringify({ type: "chat", text: text }));
    addChatMessage(text, "user");
    input.value = "";
}

// Handle Enter key
document.addEventListener("DOMContentLoaded", () => {
    const input = document.getElementById("chatInput");
    if (input) {
        input.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                sendChat();
            }
        });
    }
});
```

**Step 5:** Update the WebSocket `onmessage` handler to dispatch incoming chat messages (find the existing handler):

```javascript
ws.onmessage = async (event) => {
    const msg = JSON.parse(event.data);

    // Chat messages from the harness
    if (msg.type === "chat") {
        addChatMessage(msg.text, "agent");
        return;
    }

    // Existing command dispatch (keep this as-is)
    log(`← ${msg.command}(${JSON.stringify(msg.params).substring(0, 80)})`, "cmd");
    // ... rest of existing handler
};
```

**Step 6:** Enable the chat input in `ws.onopen`:

```javascript
ws.onopen = () => {
    setStatus("Connected", "ok");
    log("Connected to relay server", "ok");
    document.getElementById("runTest").disabled = false;
    document.getElementById("chatInput").disabled = false;
    document.getElementById("chatSend").disabled = false;
    ws.send(JSON.stringify({ type: "hello", client: "officejs-addin" }));
};
```

And disable in `ws.onclose`:

```javascript
ws.onclose = () => {
    setStatus("Disconnected — reconnecting...", "err");
    log("Disconnected, retrying in 3s...", "err");
    document.getElementById("chatInput").disabled = true;
    document.getElementById("chatSend").disabled = true;
    setTimeout(connect, 3000);
};
```

**Step 7:** Manual verification. Start the bridge server and sideload:

```bash
python3 run_bridge.py &
sleep 2
npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml
```

Open the taskpane in Excel. Confirm: chat input box visible, "Send" button enabled once connected, typing a message and hitting Enter makes the message appear in the chat log as a blue right-aligned bubble. The old "Run Test" button still works.

**Step 8:** Test the round-trip manually. With the bridge server still running, in a separate Python shell:

```python
import asyncio
from bridge_server import BridgeServer
# ...actually no, the bridge server is already running in a different process.
# Instead, write a tiny helper script.
```

Create `scripts/phase2_echo.py`:

```python
"""Phase 2 chat echo: drain chat queue and echo back to the user.

Requires the bridge server to already be running (via run_bridge.py).
"""
import asyncio
import json
import ssl
import websockets


async def main():
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    # Connect to the WSS as a second client. Won't work — the relay only accepts
    # one add-in connection. Instead, we run this test differently: the smoke test
    # needs the harness process itself to drain and send, so we'll defer this
    # to after Phase 3 when the harness exists.
    print("This script can only run from inside the harness process in later phases.")
    print("For Phase 2, verify manually:")
    print("1. Type a message in the taskpane")
    print("2. Observe it appears as a blue user bubble")
    print("3. (Phase 3+) The harness will echo it back as an agent bubble")


if __name__ == "__main__":
    asyncio.run(main())
```

Actually, drop the echo script and just test in-process.

**Step 9:** Write an in-process chat round-trip test. Add to `tests/test_bridge_server.py`:

```python
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
```

Run:

```bash
python3 -m pytest tests/test_bridge_server.py -v 2>&1
```

Expected: all tests PASS.

**Step 10:** Commit.

```bash
git add officejs-prototype/addin/taskpane.html tests/test_bridge_server.py
git commit -m "Add chat UI to taskpane with WSS round-trip

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 2 done.** Chat channel is fully plumbed; the harness can drain user messages and reply without any agent involvement.

---

## Phase 3: Planner + Q&A loop

Goal: Write the new merged Planner agent prompt, the strict `model_spec.json` JSON Schema + validator, and the Python code in `harness_v3.py` that runs the planning phase. Output: a validated spec file in `runs/<session>/model_spec.json`. No Excel writing yet.

### Task 3.1: Define the JSON Schema for `model_spec.json`

**Files:**
- Create: `schemas/model_spec.schema.json`
- Create: `tests/test_spec_schema.py`

**Step 1:** Write the failing test. `tests/test_spec_schema.py`:

```python
"""Tests for the strict Planner output schema."""
import json
from pathlib import Path

import pytest
from jsonschema import validate, ValidationError


SCHEMA_PATH = Path(__file__).parent.parent / "schemas" / "model_spec.schema.json"


@pytest.fixture
def schema():
    return json.loads(SCHEMA_PATH.read_text())


def test_valid_spec_passes(schema):
    spec = {
        "intent": "Build a 3-statement model for ConfigHub to answer what valuation the next round supports.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Centralized user-editable assumptions.",
                "must_contain": ["discount rate", "growth rates by year", "headcount plan"],
                "data_sources": ["/path/to/brief.pdf"],
            },
            {
                "name": "Revenue",
                "purpose": "Forward revenue build by product line.",
                "must_contain": ["product line breakout", "growth rate application"],
                "data_sources": ["Inputs"],
            },
        ],
        "constraints": ["fiscal year ends December 31", "USD only"],
        "out_of_scope": ["pre-2023 historicals", "monthly granularity"],
    }
    validate(instance=spec, schema=schema)


def test_cell_addresses_rejected(schema):
    spec = {
        "intent": "Build a model.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Assumptions.",
                "must_contain": ["discount rate in B7"],
                "data_sources": [],
            }
        ],
        "constraints": [],
        "out_of_scope": [],
    }
    # Schema doesn't reject this by content — but a linter step should.
    # The schema just prevents structural overprescription (no formulas/layout fields).
    # Keeping this test here as a placeholder; actual cell-address filtering is in
    # the self-review prompt, not the schema.
    validate(instance=spec, schema=schema)  # passes schema


def test_extra_fields_rejected(schema):
    """Overprescription via extra fields should fail validation."""
    spec = {
        "intent": "Build a model.",
        "sheets": [
            {
                "name": "Inputs",
                "purpose": "Assumptions.",
                "must_contain": [],
                "data_sources": [],
                "formulas": [{"address": "B7", "formula": "=0.08"}],  # <-- forbidden
            }
        ],
        "constraints": [],
        "out_of_scope": [],
    }
    with pytest.raises(ValidationError):
        validate(instance=spec, schema=schema)


def test_missing_required_fields_rejected(schema):
    spec = {"intent": "Build a model."}
    with pytest.raises(ValidationError):
        validate(instance=spec, schema=schema)
```

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_spec_schema.py -v 2>&1`

Expected: FAIL — schema file doesn't exist.

**Step 3:** Write the schema. `schemas/model_spec.schema.json`:

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "ModelSpec",
  "type": "object",
  "additionalProperties": false,
  "required": ["intent", "sheets", "constraints", "out_of_scope"],
  "properties": {
    "intent": {
      "type": "string",
      "description": "One paragraph describing what this model answers for the user.",
      "minLength": 10
    },
    "sheets": {
      "type": "array",
      "minItems": 1,
      "items": {
        "type": "object",
        "additionalProperties": false,
        "required": ["name", "purpose", "must_contain", "data_sources"],
        "properties": {
          "name": {
            "type": "string",
            "minLength": 1,
            "maxLength": 31
          },
          "purpose": {
            "type": "string",
            "minLength": 1,
            "description": "One sentence describing what this sheet is for."
          },
          "must_contain": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Outputs that must be present. Intent-level, not cell-level."
          },
          "data_sources": {
            "type": "array",
            "items": {"type": "string"},
            "description": "File paths or other sheet names this sheet draws from."
          }
        }
      }
    },
    "constraints": {
      "type": "array",
      "items": {"type": "string"},
      "description": "Hard rules the Builder must respect (e.g., 'use 8% discount rate')."
    },
    "out_of_scope": {
      "type": "array",
      "items": {"type": "string"},
      "description": "Explicit non-goals."
    }
  }
}
```

**Step 4:** Create the directory and run tests:

```bash
mkdir -p schemas
# (write the schema file)
python3 -m pytest tests/test_spec_schema.py -v 2>&1
```

Expected: all tests PASS.

**Step 5:** Commit.

```bash
git add schemas/ tests/test_spec_schema.py
git commit -m "Add strict JSON Schema for model_spec.json

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 3.2: Write the merged Planner agent prompt

**Files:**
- Create: `agents/planner_v3.md`

**Step 1:** This is a prompt-writing task, not TDD. Write the file:

```markdown
# Planner (v3)

You are the Planner for ExcelHarness v3. You read a user's brief plus any attached source files and produce a validated `model_spec.json` that captures intent, not implementation.

## Your role

You are a **product manager**, not an engineer. Your job is to understand what the user wants and pin down ambiguities, then write a spec that the Builder can execute without having to guess at missing information. You do NOT make layout, formatting, or formula decisions — those belong to the Builder, who works from `conventions.md`.

## Process

You run in three passes in the same conversation:

### Pass 1: Ambiguity detection
Read the brief, any provided source files, and `conventions.md`. Identify every decision you would have to guess at to produce a complete spec. For each ambiguity, write a question. Output a JSON array to the tool `emit_questions`:

```json
[
  {
    "id": "option_pool_timing",
    "question": "Is the option pool pre-money or post-money dilution?",
    "context": "Brief mentions '10% option pool' without specifying timing.",
    "choices": ["Pre-money (dilutes existing shareholders)", "Post-money", "No option pool"]
  }
]
```

**Rules for Pass 1:**
- Only ask questions you genuinely cannot answer from the brief or attached files.
- Prefer multiple-choice (`choices`) over open-ended when possible.
- If the brief is fully specified, return `[]` (empty array).
- Do NOT start writing the spec yet.

The user will answer each question. The answers will arrive as a new user turn in the format:
```
CLARIFICATIONS:
- option_pool_timing: Pre-money
- ...
```

### Pass 2: Spec generation
With the answers folded into your understanding, write the full spec to `runs/<session>/model_spec.json`. The spec MUST conform to `schemas/model_spec.schema.json`. The schema is strict and additive properties will be rejected.

**Rules for Pass 2:**
- `intent`: one paragraph, what the model answers for the user.
- Each sheet's `purpose` is ONE sentence.
- `must_contain` is a bulleted list of OUTPUT-level items (e.g., "gross margin by year"), not cell-level details (not "B7 = 0.38").
- `constraints` are hard business rules, not stylistic preferences (those live in `conventions.md`).
- `out_of_scope` is your commitment to the user about what you are NOT building. Use it to resolve ambiguity by narrowing scope.
- Do NOT include cell addresses, formula text, number format strings, chart types, column widths, or any implementation detail. The schema has no fields for these.

### Pass 3: Self-review
Re-read the spec you just wrote. In this pass, you are a skeptical reviewer looking for:
- **Contradictions**: Does any sheet's `must_contain` conflict with another sheet's purpose or a constraint?
- **Overprescription**: Did implementation details leak in as prose? (e.g., `must_contain: ["use VLOOKUP from Inputs!B7"]` — rewrite as "reference assumptions from Inputs").
- **Incompleteness**: Is there a sheet implied by the intent that isn't listed?
- **Redundancy**: Do two sheets overlap enough to merge?

If you find issues, rewrite the spec in place and save it again. Announce what you changed.

## What you have access to

- `Read`, `Glob`, `Grep` tools for reading the brief, source files, and `conventions.md`.
- `Write(runs/*/model_spec.json)` to save the spec.
- A tool to emit clarification questions (the harness handles the chat round-trip).

## What you do NOT do

- You do NOT write Python code.
- You do NOT decide cell addresses, formulas, or formatting.
- You do NOT invent data — if the brief is missing info, ask in Pass 1.
- You do NOT proceed to Pass 2 before clarifications arrive.

## Output expectations

At the end of your run, there is exactly one artifact: `runs/<session>/model_spec.json`, validated against the schema. The harness will hand this to the Builder.
```

**Step 2:** Verify the file exists:

Run: `wc -l agents/planner_v3.md`

Expected: ~80 lines.

**Step 3:** Commit.

```bash
git add agents/planner_v3.md
git commit -m "Add v3 Planner prompt with strict spec schema and Q&A pass

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 3.3: Write tests for the `Session` helper that owns per-run state

**Files:**
- Create: `tests/test_session.py`

**Step 1:** Write the failing test. `tests/test_session.py`:

```python
"""Tests for the Session helper that manages per-run state (paths, artifacts)."""
import json
from pathlib import Path

from session import Session


def test_session_creates_directory_structure(tmp_path):
    s = Session(root=tmp_path)
    assert s.run_dir.exists()
    assert s.run_dir.parent == tmp_path / "runs"
    assert s.snapshots_dir.exists()
    assert s.screenshots_dir.exists()


def test_session_writes_brief(tmp_path):
    s = Session(root=tmp_path)
    s.save_brief("build me a model")
    assert (s.run_dir / "brief.md").read_text() == "build me a model"


def test_session_appends_clarifications(tmp_path):
    s = Session(root=tmp_path)
    s.append_clarification("discount_rate", "8%")
    s.append_clarification("fiscal_year_end", "Dec 31")
    body = (s.run_dir / "clarifications.md").read_text()
    assert "discount_rate: 8%" in body
    assert "fiscal_year_end: Dec 31" in body


def test_session_saves_spec(tmp_path):
    s = Session(root=tmp_path)
    spec = {"intent": "x", "sheets": [], "constraints": [], "out_of_scope": []}
    s.save_spec(spec)
    loaded = json.loads((s.run_dir / "model_spec.json").read_text())
    assert loaded == spec


def test_session_append_chat_log(tmp_path):
    s = Session(root=tmp_path)
    s.append_chat("user", "hi")
    s.append_chat("agent", "hello")
    lines = (s.run_dir / "chat_log.jsonl").read_text().strip().split("\n")
    assert len(lines) == 2
    assert json.loads(lines[0])["role"] == "user"
    assert json.loads(lines[0])["text"] == "hi"
```

**Step 2:** Run:

Run: `python3 -m pytest tests/test_session.py -v 2>&1`

Expected: FAIL — no `session` module.

### Task 3.4: Implement `session.py`

**Files:**
- Create: `session.py`

**Step 1:** Implementation:

```python
"""Per-run session state: paths, artifacts, chat log."""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path


class Session:
    def __init__(self, root: Path, timestamp: str | None = None):
        self.root = Path(root)
        self.timestamp = timestamp or datetime.now().strftime("%Y%m%d-%H%M%S")
        self.run_dir = self.root / "runs" / self.timestamp
        self.snapshots_dir = self.run_dir / "snapshots"
        self.screenshots_dir = self.run_dir / "screenshots"
        for d in (self.run_dir, self.snapshots_dir, self.screenshots_dir):
            d.mkdir(parents=True, exist_ok=True)

    def save_brief(self, text: str) -> None:
        (self.run_dir / "brief.md").write_text(text)

    def append_clarification(self, question_id: str, answer: str) -> None:
        path = self.run_dir / "clarifications.md"
        with path.open("a") as f:
            f.write(f"- {question_id}: {answer}\n")

    def save_spec(self, spec: dict) -> None:
        (self.run_dir / "model_spec.json").write_text(json.dumps(spec, indent=2))

    def load_spec(self) -> dict:
        return json.loads((self.run_dir / "model_spec.json").read_text())

    def append_chat(self, role: str, text: str) -> None:
        with (self.run_dir / "chat_log.jsonl").open("a") as f:
            f.write(json.dumps({
                "ts": datetime.now().isoformat(),
                "role": role,
                "text": text,
            }) + "\n")
```

**Step 2:** Run tests:

Run: `python3 -m pytest tests/test_session.py -v 2>&1`

Expected: all PASS.

**Step 3:** Commit.

```bash
git add session.py tests/test_session.py
git commit -m "Add Session helper for per-run state

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 3.5: Write the harness_v3.py Planner phase

This task builds the harness entry point and the planning phase only. Builder/Evaluator integration comes in later phases.

**Files:**
- Create: `harness_v3.py`

**Step 1:** This is integration code. No unit test — we verify via manual run. Create `harness_v3.py`:

```python
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
```

**Step 2:** Manual verification.

```bash
# Start harness
python3 harness_v3.py &

# Sideload add-in
sleep 2
npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml
```

Expected flow:
1. Harness prints `Open Excel with the sideloaded add-in now.`
2. Taskpane opens, shows "Connected"
3. Chat pane shows: "Hi. What would you like to build?..."
4. You type a brief in the chat input, hit Enter.
5. Chat shows: "Thanks. Starting planning..."
6. Chat shows: "Spec complete: 1 sheet(s). Saved to ..."
7. Harness prints `[harness] Spec saved. Planner phase done.`
8. A file exists at `runs/<timestamp>/model_spec.json`.

Verify:

Run: `ls runs/ && cat runs/*/model_spec.json | tail -20`

Expected: the placeholder spec JSON.

**Step 3:** Stop the harness:

Run: `kill $(lsof -ti :3000) 2>&1`

**Step 4:** Commit.

```bash
git add harness_v3.py
git commit -m "Add harness_v3.py with bridge + chat plumbing, Planner stubbed

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 3.6: Wire the real Planner agent into harness_v3.py

**Files:**
- Modify: `harness_v3.py`

**Step 1:** Replace the `run_planner` stub with a real Claude Agent SDK invocation. The v2 harness has a `run_agent` helper at lines 264-345 — read it for reference but DO NOT call it (we're not depending on v2 code):

Run: `head -345 harness.py | tail -90`

**Step 2:** Add a lightweight agent runner to `harness_v3.py`. Near the top, add imports:

```python
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
```

Add a helper function:

```python
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
    # it). We'll upgrade to a session-based client in Task 3.7 if single-turn proves
    # insufficient.
    combined = "\n\n---\n\n".join(user_messages)

    final_text = ""
    async for message in query(prompt=combined, options=options):
        if isinstance(message, AssistantMessage):
            for block in message.content:
                if isinstance(block, TextBlock):
                    final_text = block.text
    return final_text
```

**Step 3:** Replace `run_planner` with the real implementation:

```python
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

    # --- Load, validate, optionally pass-3 self-review ---
    spec_path = session.run_dir / "model_spec.json"
    if not spec_path.exists():
        raise RuntimeError("Planner did not write model_spec.json")

    spec = json.loads(spec_path.read_text())
    try:
        jsonschema.validate(instance=spec, schema=schema)
    except jsonschema.ValidationError as e:
        await server.send_chat(f"Spec failed validation: {e.message}. Retrying...")
        # Let the agent retry once with the error message
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
```

Update the call site in `main()`:

```python
spec = await run_planner(session, server, brief)
```

**Step 4:** Manual verification. Start harness, provide a brief that intentionally has an ambiguity:

```bash
python3 harness_v3.py &
sleep 2
npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml
```

In the chat, type: "Build me a simple valuation model with a 10% option pool and a 20% discount rate."

Expected:
1. Planner Pass 1 asks: "Is the option pool pre-money or post-money?" (plus maybe other ambiguities)
2. You answer "Pre-money"
3. Planner Pass 2 writes `runs/<ts>/model_spec.json`
4. Spec validates against schema
5. Pass 3 runs, either returns "CLEAN" or rewrites
6. Chat says "Spec complete: N sheet(s). Saved to ..."

Inspect the spec:

Run: `cat runs/*/model_spec.json`

Expected: valid JSON matching the schema, no cell addresses or formula details.

**Step 5:** Commit.

```bash
git add harness_v3.py
git commit -m "Wire real Planner agent (3 passes, Q&A, schema validation)

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 3 done.** Planner runs end-to-end, produces a validated spec.

---

## Phase 4: Builder loop

Goal: Long-running Builder agent that writes Python scripts, runs them against the bridge, reads chat messages between turns, and calls `bridge.checkpoint()` at natural stopping points. The Evaluator is **stubbed** in this phase (returns PASS immediately) to validate the flow shape first.

### Task 4.1: Add checkpoint HTTP endpoint and queue to BridgeServer

**Files:**
- Modify: `tests/test_bridge_server.py`
- Modify: `bridge_server.py`

**Step 1:** Write the failing test. Append to `tests/test_bridge_server.py`:

```python
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
```

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_bridge_server.py::test_checkpoint_long_poll -v 2>&1`

Expected: FAIL — no `/api/checkpoint` route, no `pop_pending_checkpoint`.

**Step 3:** Implement. Add to `bridge_server.py`:

```python
from dataclasses import dataclass


@dataclass
class PendingCheckpoint:
    description: str
    future: asyncio.Future

    def resolve(self, result: dict) -> None:
        if not self.future.done():
            self.future.set_result(result)
```

In `BridgeServer.__init__`:

```python
self._checkpoint_queue: asyncio.Queue[PendingCheckpoint] = asyncio.Queue()
```

In `start()`, add the route:

```python
app.router.add_post("/api/checkpoint", self._handle_checkpoint)
```

Add handlers:

```python
async def _handle_checkpoint(self, request: web.Request) -> web.Response:
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON"}, status=400)
    description = body.get("description", "")
    future = asyncio.get_event_loop().create_future()
    pending = PendingCheckpoint(description=description, future=future)
    await self._checkpoint_queue.put(pending)
    result = await future
    return web.json_response(result)

def pop_pending_checkpoint(self) -> PendingCheckpoint | None:
    try:
        return self._checkpoint_queue.get_nowait()
    except asyncio.QueueEmpty:
        return None

async def wait_for_checkpoint(self) -> PendingCheckpoint:
    return await self._checkpoint_queue.get()
```

**Step 4:** Run tests:

Run: `python3 -m pytest tests/test_bridge_server.py -v 2>&1`

Expected: all PASS.

**Step 5:** Add `checkpoint()` to `bridge.py`:

```python
def checkpoint(self, description: str) -> dict:
    """Block until the harness resolves this checkpoint (after evaluator + commit)."""
    resp = requests.post(
        f"{self.base_url}/api/checkpoint",
        json={"description": description},
        verify=self.verify,
        timeout=600,  # long poll — checkpoint may take a while
    )
    resp.raise_for_status()
    return resp.json()
```

Also add `emit()` for agent-initiated chat:

```python
# Note: chat messages from Builder scripts need a new HTTP route.
# We'll defer the `emit()` method to Task 4.2.
```

Actually add the route and method now:

In `bridge_server.py`, add route in `start()`:

```python
app.router.add_post("/api/emit", self._handle_emit)
```

Handler:

```python
async def _handle_emit(self, request: web.Request) -> web.Response:
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON"}, status=400)
    text = body.get("text", "")
    await self.send_chat(text)
    return web.json_response({"ok": True})
```

In `bridge.py`:

```python
def emit(self, text: str) -> dict:
    """Send an agent-initiated chat message to the taskpane."""
    resp = requests.post(
        f"{self.base_url}/api/emit",
        json={"text": text},
        verify=self.verify,
        timeout=self.timeout,
    )
    resp.raise_for_status()
    return resp.json()
```

**Step 6:** Commit.

```bash
git add bridge_server.py bridge.py tests/test_bridge_server.py
git commit -m "Add checkpoint and emit HTTP endpoints to bridge

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 4.2: Write the v3 Builder agent prompt

**Files:**
- Create: `agents/builder_v3.md`

```markdown
# Builder (v3)

You are the Builder for ExcelHarness v3. You read the validated spec at `runs/<session>/model_spec.json` and build the model live inside Excel by writing Python scripts that use `bridge.py`.

## Your tools

- `Read`, `Glob`, `Grep` for inspecting the spec, `conventions.md`, and input files.
- `Write(/tmp/builder_*.py)` to write build scripts into a temp directory.
- `Bash(python3 /tmp/builder_*.py)` to execute them.
- `Bash(ls*)`, `Bash(cat*)` for shell inspection.

You do NOT have access to `openpyxl`, `pandas`, or any Excel library. Your only way to affect the workbook is through `bridge.py`.

## Your operating loop

You run as a **long-running agent**. The user is watching the workbook update live as your scripts run. Between script runs, the harness injects any new chat messages from the user as your next turn.

Each turn, you:

1. **Read context** — the spec, `conventions.md`, and any prior build state. Use `bridge.dump_sheet()` to see what's already in the workbook.
2. **Plan the next unit of work** — one cohesive chunk (a section, a sheet, a set of related formulas). Do NOT try to build the whole model in one script.
3. **Write a Python script** to `/tmp/builder_NN.py` (pick a monotonically increasing number).
4. **Run it** with Bash.
5. **Check the output** for errors. Use `bridge.read_values()` or `bridge.dump_sheet()` to verify what you built.
6. **Emit a chat message** via `bridge.emit()` at important moments ("Starting Revenue sheet", "Completed calculations").
7. **At natural stopping points**, call `bridge.checkpoint("short description")`. This blocks until the Evaluator has reviewed your work.
    - If the checkpoint returns `{status: "pass"}`: the harness committed to git; continue.
    - If it returns `{status: "fail", findings: [...]}`: read findings, fix the issues in your next script, then retry the checkpoint.

## Rules

- **Every script is idempotent** if possible. If the harness crashes mid-build, your next turn should be able to read the workbook and resume. Use `bridge.dump_sheet()` to figure out where you are.
- **Follow conventions.md rigorously.** Blue font for hardcoded inputs, black for calculations, green for cross-sheet refs. Explicit formatting everywhere.
- **Check your work.** After writing formulas, use `bridge.read_values()` to verify they produce sensible numbers. If a formula evaluates to `#REF!` or a number that's off by 1000x, fix it before moving on.
- **Read chat messages.** The harness will inject chat messages from the user as user turns. Treat them as directives: "make column C wider" means you should widen it in your next script.
- **Checkpoint often enough to commit meaningful progress** (per sheet, or per major section within a large sheet), but not so often that git history becomes noise.

## Starting a session

Your first turn will have the spec and a brief instruction. Start by reading the spec, `conventions.md`, and any files listed in `sheets[].data_sources`. Then begin building the first sheet.

## Ending a session

When you've built everything in the spec and all checkpoints have passed, emit a final chat message: "Model complete. Ready for review." Then your turn ends with no further tool calls. The harness will take over and unprotect the workbook.

## Example script shape

```python
#!/usr/bin/env python3
"""Build Revenue sheet — product line breakout by year."""
from bridge import Bridge, BridgeError

b = Bridge(base_url="https://localhost:3000", verify_tls=False)

b.create_sheet("Revenue")
b.write_values("Revenue", "A1:F1", [["Line", "2024", "2025", "2026", "2027", "Total"]])
b.format_range("Revenue", "A1:F1", {
    "font": {"bold": True, "color": "#FFFFFF"},
    "fill": {"color": "#2F5496"},
    "horizontalAlignment": "Center",
    "borders": {"bottom": {"style": "Continuous", "color": "#1F3864", "weight": "Thick"}},
})
# ... more build commands
b.emit("Revenue sheet structure complete, applying formulas next")
```
```

**Commit:**

```bash
git add agents/builder_v3.md
git commit -m "Add v3 Builder prompt for long-running live-Excel agent

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 4.3: Implement the Builder loop in harness_v3.py

**Files:**
- Modify: `harness_v3.py`

**Step 1:** Add a helper class to track builder state. Add after `run_planner`:

```python
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

    # We'll use the streaming query API. The loop structure:
    #   - Run one agent turn (streaming).
    #   - When turn ends, drain chat queue, drain checkpoint queue.
    #   - If checkpoint pending: handle it (snapshot + stub evaluator + commit + resolve).
    #   - If chat pending: inject as next user message.
    #   - Otherwise: inject "continue" as next user message.
    #   - Stop when agent says "Model complete. Ready for review."

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
            while not stop_watcher.is_set():
                pending_cp = server.pop_pending_checkpoint()
                if pending_cp is not None:
                    # Snapshot the workbook
                    snap_result = await server.send_command("saveSnapshot", {})
                    if snap_result.get("ok"):
                        import base64
                        xlsx_bytes = base64.b64decode(snap_result["base64"])
                        snap_path = session.snapshots_dir / f"turn_{turn}.xlsx"
                        snap_path.write_bytes(xlsx_bytes)
                    # Stub: evaluator always passes in Phase 4
                    verdict = {"status": "pass"}
                    # Git commit
                    model_path = session.run_dir / "models"
                    model_path.mkdir(exist_ok=True)
                    committed = model_path / "model.xlsx"
                    committed.write_bytes(xlsx_bytes)
                    import subprocess
                    subprocess.run(["git", "add", str(committed)], check=False, cwd=ROOT)
                    subprocess.run(
                        ["git", "commit", "-m", f"Checkpoint: {pending_cp.description}"],
                        check=False,
                        cwd=ROOT,
                    )
                    await server.send_chat(f"✓ Committed: {pending_cp.description}")
                    pending_cp.resolve(verdict)
                    checkpoint_results.append(pending_cp.description)
                await asyncio.sleep(0.3)

        watcher_task = asyncio.create_task(watch_checkpoints())

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
```

Update `main()` to call it:

```python
await run_builder_loop(session, server, spec)
```

**Step 2:** Manual verification.

```bash
python3 harness_v3.py &
sleep 2
npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml
```

Type a simple brief in the chat: "Build me a tiny spreadsheet: one sheet called 'Test' with a header row 'A', 'B', 'C' and three rows of fake data."

Expected:
1. Planner runs (maybe asks no questions), spec saved.
2. Builder starts. Taskpane shows chat messages from the Builder.
3. Cells appear in Excel.
4. At some point the Builder calls `bridge.checkpoint()` and the harness logs "✓ Committed: ..." to chat.
5. `git log --oneline` shows a new checkpoint commit.
6. Builder eventually says "Model complete. Ready for review."
7. Workbook becomes editable.

**Step 3:** Commit.

```bash
git add harness_v3.py
git commit -m "Add Builder loop with stubbed evaluator and checkpoint handling

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 4 done.** Builder runs long-form, checkpoints, commits, and the flow shape is validated with a stubbed evaluator.

---

## Phase 5: Evaluator

Goal: Replace the stubbed evaluator in the checkpoint handler with a real adversarial Evaluator agent. It uses `dump_sheet` on each sheet the Builder claims to have completed, plus rendered PNGs from LibreOffice.

### Task 5.1: Write the LibreOffice snapshot renderer

**Files:**
- Create: `snapshot_renderer.py`
- Create: `tests/test_snapshot_renderer.py`

**Step 1:** Write the failing test. `tests/test_snapshot_renderer.py`:

```python
"""Tests for LibreOffice-based xlsx -> per-sheet PNG renderer."""
from pathlib import Path

import pytest

from snapshot_renderer import render_xlsx_to_pngs


def test_renders_pngs_for_each_sheet(tmp_path):
    """Given a fixture xlsx, render should produce one PNG per sheet."""
    fixture = Path(__file__).parent / "fixtures" / "two_sheet.xlsx"
    if not fixture.exists():
        pytest.skip("fixture not present — create via the bridge in a real session")
    out_dir = tmp_path / "screenshots"
    pngs = render_xlsx_to_pngs(fixture, out_dir)
    assert len(pngs) >= 1
    for p in pngs:
        assert p.exists()
        assert p.suffix == ".png"
        assert p.stat().st_size > 1000  # non-empty image
```

**Step 2:** Run it:

Run: `python3 -m pytest tests/test_snapshot_renderer.py -v 2>&1`

Expected: either FAIL (module doesn't exist) or SKIP (fixture missing). Either way, next step.

**Step 3:** Write the implementation. `snapshot_renderer.py`:

```python
"""Render an .xlsx into per-sheet PNGs via LibreOffice + pdftoppm.

Used by the Evaluator to get visual feedback on the live workbook's state
at checkpoint time. Office.js has no worksheet image export, so we go
via a saved snapshot + LibreOffice.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


def render_xlsx_to_pngs(xlsx_path: Path, out_dir: Path) -> list[Path]:
    """Convert xlsx -> PDF -> per-sheet PNGs.

    Returns the list of produced PNG paths, in page order.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: xlsx -> PDF via LibreOffice headless
    pdf_path = out_dir / f"{xlsx_path.stem}.pdf"
    result = subprocess.run(
        [
            "soffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(xlsx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    # Step 2: PDF -> PNGs via pdftoppm
    png_prefix = out_dir / f"{xlsx_path.stem}_page"
    result = subprocess.run(
        ["pdftoppm", "-png", "-r", "150", str(pdf_path), str(png_prefix)],
        capture_output=True,
        text=True,
        timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr}")

    pngs = sorted(out_dir.glob(f"{xlsx_path.stem}_page-*.png"))
    return pngs


def ensure_tools_available() -> None:
    """Raise if soffice or pdftoppm is not on PATH."""
    for tool in ("soffice", "pdftoppm"):
        if shutil.which(tool) is None:
            raise RuntimeError(
                f"{tool} not found on PATH. "
                "Install LibreOffice (brew install --cask libreoffice) "
                "and poppler (brew install poppler)."
            )
```

**Step 4:** Create a fixture by running an earlier build and copying the result.

```bash
mkdir -p tests/fixtures
# You should have a snapshot from a Phase 4 run somewhere in runs/*/snapshots/
cp runs/*/snapshots/*.xlsx tests/fixtures/two_sheet.xlsx 2>/dev/null || \
  echo "No snapshot yet — skipping fixture copy, test will SKIP"
```

Run:

```bash
python3 -m pytest tests/test_snapshot_renderer.py -v 2>&1
```

Expected: PASS if fixture exists, SKIP otherwise.

**Step 5:** Commit.

```bash
git add snapshot_renderer.py tests/test_snapshot_renderer.py tests/fixtures/ 2>/dev/null
git commit -m "Add LibreOffice-based snapshot renderer for Evaluator

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 5.2: Write the v3 Evaluator agent prompt

**Files:**
- Create: `agents/evaluator_v3.md`

```markdown
# Evaluator (v3)

You are the adversarial Evaluator for ExcelHarness v3. The Builder has completed a checkpoint ("Revenue sheet complete", "Balance Sheet wired up", etc.) and has asked for your review. You read the spec, the current live workbook state via `dumpSheet`, and rendered PNG screenshots, then return PASS or FAIL with specific findings.

## Your stance

You are **adversarial**. The Builder is competent but not perfect. Your job is to find real problems before the user does:

- Missing required outputs from the spec (`must_contain` items)
- Formulas that reference the wrong cells or produce nonsensical values
- Cells that should be formulas but are hardcoded (or vice versa)
- Convention violations (hardcoded inputs in black instead of blue; cross-sheet refs in black instead of green; no explicit formatting)
- Broken cross-sheet references
- Visual problems visible in the screenshots (misaligned headers, missing number formats, columns too narrow to show numbers)
- Math that doesn't balance (balance sheet that doesn't balance, cash flow that doesn't tie out)

You are NOT here to suggest improvements. Only flag concrete problems.

## Your tools

- `Read` for the spec, `conventions.md`, and screenshot PNGs.
- `Glob`, `Grep` for finding artifacts.
- `Bash` for reading structured dump output.

You do NOT have network access, cannot call the bridge directly. The harness will have written the relevant dumps and screenshots to `runs/<session>/eval_input/` before invoking you.

## Input format

The harness writes these files for you to review:

```
runs/<session>/eval_input/
├── checkpoint.md          # The description of the checkpoint you're reviewing
├── spec.json              # Copy of the spec
├── dumps/
│   ├── Inputs.json        # dumpSheet output for each sheet
│   ├── Revenue.json
│   └── ...
└── screenshots/
    ├── page-01.png
    ├── page-02.png
    └── ...
```

## Your output

Return a JSON object to stdout in the format:

```json
{
  "status": "pass" | "fail",
  "findings": [
    {
      "severity": "error" | "warning",
      "sheet": "Revenue",
      "cell": "B7",
      "issue": "Cell B7 is hardcoded at 1000 but should reference Inputs!B4 (growth rate)."
    }
  ]
}
```

- If status is `"pass"`, `findings` should be an empty array.
- If ANY finding has severity `"error"`, status MUST be `"fail"`.
- Warnings alone do not fail a checkpoint — they're surfaced to the user but don't block.
- Be specific. "Revenue doesn't look right" is useless. "Revenue B7 = $1.2B, but Inputs!B4 shows growth rate of 5% and B6 shows base of $1M — expected $1.05M." is useful.

## Review process

1. Read `checkpoint.md` to understand what scope to review. Do not review sheets the Builder hasn't touched yet.
2. Read `spec.json` to understand the intent and requirements.
3. Read `conventions.md` to understand the formatting rules.
4. For each sheet the checkpoint covers:
    a. Read `dumps/<Sheet>.json` (structured formulas + values + formatting).
    b. Look at the corresponding screenshot.
    c. Check spec compliance, math correctness, convention adherence.
5. Aggregate findings and return the JSON verdict.

## What you do NOT do

- Do not suggest stylistic improvements.
- Do not rewrite formulas.
- Do not modify any files.
- Do not run the Builder's scripts.
- Do not judge sheets that aren't in scope for this checkpoint.
```

**Commit:**

```bash
git add agents/evaluator_v3.md
git commit -m "Add v3 Evaluator prompt for adversarial live-workbook review

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 5.3: Wire the Evaluator into the checkpoint handler

**Files:**
- Modify: `harness_v3.py`

**Step 1:** Write a helper that prepares the eval input directory:

```python
async def prepare_eval_input(session: Session, server: BridgeServer, checkpoint_description: str, spec: dict, turn: int) -> Path:
    """Snapshot the workbook, dump each sheet, render screenshots, and write them into an eval_input dir."""
    from snapshot_renderer import render_xlsx_to_pngs
    import base64

    eval_dir = session.run_dir / "eval_input"
    # Clean and recreate each checkpoint so the evaluator sees only current state.
    import shutil
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
    xlsx_bytes = base64.b64decode(snap_result["base64"])
    snap_path = session.snapshots_dir / f"turn_{turn}.xlsx"
    snap_path.write_bytes(xlsx_bytes)

    # Render PNGs
    render_xlsx_to_pngs(snap_path, eval_dir / "screenshots")

    # Dump each sheet
    sheets_result = await server.send_command("getSheetNames", {})
    for sheet_name in sheets_result.get("sheets", []):
        dump = await server.send_command("dumpSheet", {"sheet": sheet_name})
        safe_name = sheet_name.replace("/", "_").replace(" ", "_")
        (eval_dir / "dumps" / f"{safe_name}.json").write_text(json.dumps(dump, indent=2))

    return eval_dir


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

    # Extract JSON from response
    import re
    m = re.search(r"\{.*\}", result_text, re.DOTALL)
    if not m:
        return {"status": "fail", "findings": [{"severity": "error", "issue": "Evaluator did not return JSON"}]}
    try:
        verdict = json.loads(m.group(0))
    except json.JSONDecodeError as e:
        return {"status": "fail", "findings": [{"severity": "error", "issue": f"Evaluator JSON parse failed: {e}"}]}
    return verdict
```

**Step 2:** Update the `watch_checkpoints` inner function in `run_builder_loop` to call the real evaluator:

Replace the stub section:

```python
# Stub: evaluator always passes in Phase 4
verdict = {"status": "pass"}
```

with:

```python
# Real Evaluator
eval_dir = await prepare_eval_input(session, server, pending_cp.description, spec, turn)
verdict = await run_evaluator(session, eval_dir)
```

And only commit if PASS:

```python
if verdict.get("status") == "pass":
    # Git commit
    import subprocess
    subprocess.run(["git", "add", str(committed)], check=False, cwd=ROOT)
    subprocess.run(
        ["git", "commit", "-m", f"Checkpoint: {pending_cp.description}"],
        check=False, cwd=ROOT,
    )
    await server.send_chat(f"✓ {pending_cp.description}")
else:
    # FAIL: don't commit. Pass findings back to Builder.
    findings_text = "\n".join(
        f"  - [{f.get('severity', 'error')}] {f.get('sheet', '')}{(':' + f['cell']) if f.get('cell') else ''}: {f.get('issue', '')}"
        for f in verdict.get("findings", [])
    )
    await server.send_chat(f"✗ Evaluator FAIL: {pending_cp.description}\n{findings_text}")

pending_cp.resolve(verdict)
```

Also update the stubbed xlsx write so both paths have it:

```python
model_path = session.run_dir / "models"
model_path.mkdir(exist_ok=True)
committed = model_path / "model.xlsx"
committed.write_bytes(xlsx_bytes)  # The bytes from prepare_eval_input

if verdict.get("status") == "pass":
    # ... commit
```

(You'll need to return `xlsx_bytes` from `prepare_eval_input` or re-read the snapshot.)

**Step 3:** Manual verification. Same flow as Task 4.3 but with the real evaluator engaged. The eval_input/ directory should be populated after each checkpoint, and the evaluator should emit real findings (pass or fail).

Test the FAIL path with a contrived broken build: tell the Builder to intentionally hardcode a value that the spec says should be a formula. The evaluator should catch it and the Builder should retry.

**Step 4:** Commit.

```bash
git add harness_v3.py
git commit -m "Wire real Evaluator into checkpoint handler

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 5 done.** Evaluator runs at each checkpoint with live dumps + PNGs.

---

## Phase 6: Deletion and cutover

Goal: Delete the v2 code (`harness.py`, `dump.py`, `deterministic_evaluator.py`, old agent prompts, `scripts/`, `evals/`, `progress.txt`). Rename `harness_v3.py` to `harness.py`. Merge `v3-live-excel` to `main`.

### Task 6.1: Verify v3 has parity with v2 on one real build

**Files:** none modified.

**Step 1:** Pick a representative brief that v2 has built successfully in the past (check `runs/` for an old one with a good outcome).

**Step 2:** Run v3 end-to-end with that brief. Confirm:
- Planner produces a valid spec.
- Builder builds all the spec's sheets.
- Evaluator runs at each checkpoint and produces PASS verdicts (or FAIL that the Builder fixes).
- Final workbook opens in Excel and looks right.

**Step 3:** If there's any regression, stop and fix it before deleting anything.

**Step 4:** Commit no code, but record the verification in a note:

```bash
git commit --allow-empty -m "Verify v3 parity with v2 on <brief description>

Manual verification: ran <brief> end-to-end through harness_v3.py.
Planner spec validated. Builder completed <N> checkpoints. Evaluator
passed <N/M>. Final workbook visually correct.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 6.2: Delete v2 Python code

**Files:**
- Delete: `harness.py` (v2, 925 lines)
- Delete: `dump.py`
- Delete: `deterministic_evaluator.py`

**Step 1:**

```bash
git rm harness.py dump.py deterministic_evaluator.py
```

**Step 2:** Commit.

```bash
git commit -m "Delete v2 harness, dump.py, deterministic_evaluator.py

Replaced by harness_v3.py, bridge.py (live Excel via Office.js), and
evaluator_v3.md (live queries + rendered PNGs). See
docs/plans/2026-04-10-live-excel-architecture.md for design context.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 6.3: Delete v2 agent prompts

**Files:**
- Delete: `agents/planner.md`, `agents/planner_v2.md`, `agents/builder.md`, `agents/evaluator_v2.md`, `agents/scoper.md` (whichever exist)

**Step 1:** Inventory:

Run: `ls agents/`

**Step 2:** Delete old ones, keeping only the `_v3` files:

```bash
git rm agents/planner.md agents/planner_v2.md agents/builder.md agents/evaluator_v2.md agents/scoper.md 2>/dev/null
```

Only delete files that actually exist.

**Step 3:** Verify v3 prompts remain:

Run: `ls agents/`

Expected: `builder_v3.md`, `evaluator_v3.md`, `planner_v3.md`.

**Step 4:** Commit.

```bash
git commit -m "Delete v2 agent prompts, keep only v3 variants

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 6.4: Delete v2 artifact directories

**Files:**
- Delete: `scripts/` (except `phase1_smoke.py`, which we keep as a smoke test tool)
- Delete: `evals/` (if any)
- Delete: `progress.txt` (if exists)

**Step 1:** Check contents:

```bash
ls scripts/ 2>/dev/null
ls evals/ 2>/dev/null
ls progress.txt 2>/dev/null
```

**Step 2:** Remove old scripts (but keep the Phase 1 smoke test):

```bash
find scripts/ -type f ! -name 'phase1_smoke.py' ! -name 'phase2_*.py' -delete 2>/dev/null
rm -rf evals/ progress.txt 2>/dev/null
git add -A
```

**Step 3:** Commit.

```bash
git commit -m "Delete v2 artifact directories (scripts/, evals/, progress.txt)

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 6.5: Rename `harness_v3.py` to `harness.py` and rename agent prompts

**Files:**
- Rename: `harness_v3.py` -> `harness.py`
- Rename: `agents/planner_v3.md` -> `agents/planner.md`
- Rename: `agents/builder_v3.md` -> `agents/builder.md`
- Rename: `agents/evaluator_v3.md` -> `agents/evaluator.md`

**Step 1:** Rename files:

```bash
git mv harness_v3.py harness.py
git mv agents/planner_v3.md agents/planner.md
git mv agents/builder_v3.md agents/builder.md
git mv agents/evaluator_v3.md agents/evaluator.md
```

**Step 2:** Update internal references in `harness.py` that point at `planner_v3.md`, `builder_v3.md`, `evaluator_v3.md`:

```bash
grep -l "_v3.md" harness.py
```

Edit `harness.py` and change those three references to the new names.

**Step 3:** Run the tests to make sure nothing broke.

Run: `python3 -m pytest -v 2>&1`

Expected: all PASS.

**Step 4:** Manual smoke test: run the harness once more with a tiny brief to confirm nothing broke in the rename.

**Step 5:** Commit.

```bash
git add harness.py agents/
git commit -m "Rename v3 files to drop _v3 suffix (v2 fully removed)

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

### Task 6.6: Update README and merge to main

**Files:**
- Create or modify: `README.md`

**Step 1:** Check if README exists:

```bash
ls README.md 2>/dev/null
```

**Step 2:** Create or update with a short description. `README.md`:

```markdown
# ExcelHarness v3

Claude-driven Excel model builder that works inside a live Excel session via an Office.js add-in. The harness runs a Planner, Builder, and Evaluator against a live workbook, coordinated through a chat panel in the Excel taskpane.

## Requirements
- macOS (tested on Excel 16.107+)
- Python 3.13, `pip install aiohttp websockets requests claude-agent-sdk jsonschema`
- Microsoft Excel with the sideloaded add-in (`officejs-prototype/addin/manifest.xml`)
- LibreOffice + poppler (`brew install --cask libreoffice; brew install poppler`) for snapshot rendering

## Usage
1. Trust the dev certificate once: `sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain officejs-prototype/certs/cert.pem`
2. Start the harness: `python3 harness.py`
3. Sideload the add-in: `npx office-addin-dev-settings sideload officejs-prototype/addin/manifest.xml`
4. Once the taskpane shows "Connected", type your brief into the chat.
5. The Planner will ask any clarifying questions in chat. Answer them.
6. Watch the Builder build the model. Steer via chat as needed.
7. The Evaluator reviews at each checkpoint and commits to git on pass.

## Architecture
See [docs/plans/2026-04-10-live-excel-architecture.md](docs/plans/2026-04-10-live-excel-architecture.md).
```

**Step 3:** Commit and merge.

```bash
git add README.md
git commit -m "Add README for v3

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"

# Check v2-three-agent isn't behind:
git log --oneline v2-three-agent..v3-live-excel | wc -l
# Should be ~25+ commits

# Merge to main
git checkout main
git merge --no-ff v3-live-excel -m "Merge v3 live-Excel architecture

Replaces the v2 three-agent openpyxl pipeline with a single-process
harness that drives live Excel via an Office.js bridge. Chat is the
sole steering mechanism. See docs/plans/2026-04-10-live-excel-architecture.md.

Co-Authored-By: Claude Opus 4.6 (1M context) <noreply@anthropic.com>"
```

**Phase 6 done.** v3 is on main.

---

## Appendices

### A. Running the full test suite

```bash
python3 -m pytest -v
```

At the end of Phase 5, expect ~15-20 tests covering: bridge server lifecycle, command dispatch, HTTP endpoints, chat channel, checkpoint long-poll, chat queue, bridge client, session helper, and spec schema.

### B. Known non-blockers

- **freezeRows** fails on Mac Office.js 16.107+ with "internal error". Confirmed Office.js bug. The Builder should skip `freeze_rows` on Mac or catch the exception and continue. Not worth retrying.
- **LibreOffice render fidelity** differs slightly from Excel's native render (fonts, subtle formatting). Good enough for the Evaluator; if a specific finding looks like a render artifact, confirm in the live workbook before acting.

### C. Debugging tips

- **Taskpane not connecting:** Check cert trust (`security find-certificate -c localhost`), check bridge server is listening (`lsof -i :3000`), check manifest sideload.
- **WebView caching old taskpane:** Bump a query string on the `SourceLocation` in the manifest and re-sideload.
- **Builder scripts fail silently:** Check Bash tool output in the agent's stream; the harness doesn't currently surface script stderr to chat automatically.
- **Evaluator returns non-JSON:** Look at the full `final_text` returned by `run_agent_session` — the Evaluator prompt may need tightening if it's chattering before the JSON.

### D. What's deliberately NOT in this plan

- Migrating existing `runs/` from v2 format to v3. v2 runs are archived as-is; no conversion.
- Multi-workbook or multi-user sessions.
- Any attempt to support headless mode. Always-live is the operating model.
- A CLI for replaying a session from git history. Nice to have, not needed for the rewrite.
