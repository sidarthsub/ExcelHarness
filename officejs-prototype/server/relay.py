#!/usr/bin/env python3
"""
WebSocket relay server for Office.js add-in prototype.
Serves the add-in files over HTTPS and provides a WSS endpoint
for sending Office.js commands to the live Excel workbook.
"""

import asyncio
import json
import ssl
import os
from pathlib import Path

import websockets
from aiohttp import web

ROOT = Path(__file__).parent.parent
ADDIN_DIR = ROOT / "addin"
CERTS_DIR = ROOT / "certs"

# Track the connected add-in client
addin_connection = None
pending_responses = {}
msg_counter = 0


async def handle_addin_ws(websocket):
    """Handle the WebSocket connection from the Office.js add-in."""
    global addin_connection
    addin_connection = websocket
    print("[relay] Add-in connected")
    try:
        async for message in websocket:
            data = json.loads(message)
            msg_id = data.get("id")
            if msg_id and msg_id in pending_responses:
                pending_responses[msg_id].set_result(data)
            else:
                print(f"[relay] Add-in says: {data}")
    except websockets.exceptions.ConnectionClosed:
        print("[relay] Add-in disconnected")
    finally:
        addin_connection = None


async def send_command(command: str, params: dict = None, timeout: float = 30.0) -> dict:
    """Send a command to the add-in and wait for the response."""
    global msg_counter
    if not addin_connection:
        return {"error": "No add-in connected"}

    msg_counter += 1
    msg_id = f"cmd-{msg_counter}"
    payload = {"id": msg_id, "command": command, "params": params or {}}

    future = asyncio.get_event_loop().create_future()
    pending_responses[msg_id] = future

    await addin_connection.send(json.dumps(payload))
    try:
        result = await asyncio.wait_for(future, timeout=timeout)
        return result
    except asyncio.TimeoutError:
        pending_responses.pop(msg_id, None)
        return {"error": "timeout"}


async def handle_api(request):
    """HTTP API endpoint to send commands from external scripts."""
    try:
        body = await request.json()
    except Exception:
        return web.json_response({"error": "invalid JSON"}, status=400)

    command = body.get("command")
    params = body.get("params", {})
    if not command:
        return web.json_response({"error": "missing 'command'"}, status=400)

    result = await send_command(command, params)
    return web.json_response(result)


async def handle_static(request):
    """Serve add-in static files."""
    path = request.match_info.get("path", "taskpane.html")
    if not path or path == "/":
        path = "taskpane.html"
    file_path = ADDIN_DIR / path
    if file_path.exists() and file_path.is_file():
        resp = web.FileResponse(file_path)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    return web.Response(status=404, text="Not found")


async def run_test_sequence():
    """Run a test sequence after the add-in connects."""
    print("[test] Waiting for add-in connection...")
    while not addin_connection:
        await asyncio.sleep(0.5)
    print("[test] Add-in connected! Running test sequence...\n")

    # Test 1: Create a sheet
    print("--- Test 1: Create sheet ---")
    r = await send_command("createSheet", {"name": "Prototype Test"})
    print(f"  Result: {r}\n")

    # Test 2: Write values
    print("--- Test 2: Write values ---")
    r = await send_command("writeValues", {
        "sheet": "Prototype Test",
        "address": "A1:D1",
        "values": [["Company", "Revenue", "Growth", "Multiple"]]
    })
    print(f"  Result: {r}\n")

    # Test 3: Write data rows
    print("--- Test 3: Write data ---")
    r = await send_command("writeValues", {
        "sheet": "Prototype Test",
        "address": "A2:D6",
        "values": [
            ["ConfigHub", 5000000, 0.15, 6],
            ["Acme Corp", 12000000, 0.25, 8],
            ["Widget Inc", 3000000, 0.10, 5],
            ["DataFlow", 8000000, 0.30, 10],
            ["CloudNet", 20000000, 0.20, 7],
        ]
    })
    print(f"  Result: {r}\n")

    # Test 4: Write formulas
    print("--- Test 4: Write formulas ---")
    r = await send_command("writeFormulas", {
        "sheet": "Prototype Test",
        "address": "E1",
        "formulas": [["Valuation"]]
    })
    print(f"  Header: {r}")
    r = await send_command("writeFormulas", {
        "sheet": "Prototype Test",
        "address": "E2:E6",
        "formulas": [
            ["=B2*D2"], ["=B3*D3"], ["=B4*D4"], ["=B5*D5"], ["=B6*D6"]
        ]
    })
    print(f"  Formulas: {r}\n")

    # Test 5: Format header row
    print("--- Test 5: Format header ---")
    r = await send_command("formatRange", {
        "sheet": "Prototype Test",
        "address": "A1:E1",
        "format": {
            "font": {"bold": True, "size": 11, "color": "#FFFFFF"},
            "fill": {"color": "#2F5496"},
            "horizontalAlignment": "Center",
            "borders": {
                "bottom": {"style": "Continuous", "color": "#1F3864", "weight": "Thick"}
            }
        }
    })
    print(f"  Result: {r}\n")

    # Test 6: Number formats
    print("--- Test 6: Number formats ---")
    r = await send_command("setNumberFormat", {
        "sheet": "Prototype Test",
        "address": "B2:B6",
        "format": "$#,##0"
    })
    print(f"  Revenue: {r}")
    r = await send_command("setNumberFormat", {
        "sheet": "Prototype Test",
        "address": "C2:C6",
        "format": "0.0%"
    })
    print(f"  Growth: {r}")
    r = await send_command("setNumberFormat", {
        "sheet": "Prototype Test",
        "address": "E2:E6",
        "format": "$#,##0"
    })
    print(f"  Valuation: {r}\n")

    # Test 7: Column widths
    print("--- Test 7: Column widths ---")
    r = await send_command("setColumnWidths", {
        "sheet": "Prototype Test",
        "columns": {"A": 120, "B": 110, "C": 80, "D": 80, "E": 120}
    })
    print(f"  Result: {r}\n")

    # Test 8: Create a chart
    print("--- Test 8: Create chart ---")
    r = await send_command("createChart", {
        "sheet": "Prototype Test",
        "type": "ColumnClustered",
        "dataRange": "A1:B6",
        "seriesBy": "Columns",
        "title": "Revenue by Company",
        "top": 150, "left": 10, "width": 500, "height": 300,
        "axisFormat": {"valueAxis": {"numberFormat": "$#,##0"}},
        "seriesColors": ["#2F5496"]
    })
    print(f"  Result: {r}\n")

    # Test 9: Read back computed values
    print("--- Test 9: Read computed values ---")
    r = await send_command("readValues", {
        "sheet": "Prototype Test",
        "address": "E2:E6"
    })
    print(f"  Valuations: {r}\n")

    # Test 10: Freeze panes
    print("--- Test 10: Freeze panes ---")
    r = await send_command("freezeRows", {
        "sheet": "Prototype Test",
        "count": 1
    })
    print(f"  Result: {r}\n")

    print("=" * 50)
    print("ALL TESTS COMPLETE")
    print("=" * 50)


async def main():
    ssl_ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ssl_ctx.load_cert_chain(CERTS_DIR / "cert.pem", CERTS_DIR / "key.pem")

    # HTTPS server for static files + API
    app = web.Application()
    app.router.add_post("/api/command", handle_api)
    app.router.add_get("/{path:.*}", handle_static)

    runner = web.AppRunner(app)
    await runner.setup()
    https_site = web.TCPSite(runner, "localhost", 3000, ssl_context=ssl_ctx)
    await https_site.start()
    print("[relay] HTTPS server on https://localhost:3000")

    # WSS server
    wss_server = await websockets.serve(
        handle_addin_ws, "localhost", 3001, ssl=ssl_ctx
    )
    print("[relay] WSS server on wss://localhost:3001")
    print("[relay] Waiting for add-in to connect...\n")

    await asyncio.Future()  # run forever


if __name__ == "__main__":
    asyncio.run(main())
