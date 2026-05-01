"""Pseudo-headless bridge server.

Drop-in replacement for the live Office.js bridge (bridge_server.py). Listens
on the same HTTP endpoints (`/api/command`, `/api/checkpoint`, `/api/emit`)
but routes commands to a real hidden Excel instance via xlwings instead of
over WebSocket to the Office.js add-in.

Why xlwings: the downstream pipeline (snapshot_renderer, live_dump, the
evaluator) expects xlsx files produced by real Excel — cached formula
values, proper format tables, iterative-calc cache. Openpyxl-authored
files break those. xlwings drives real Excel hidden in the background.

Usage from Python:

    with PseudoBridgeServer(
        output_xlsx="/path/to/model.xlsx",
        port=3100,
    ) as server:
        # spawn builder here — builder's Bridge() hits http://localhost:3100
        ...
    # on exit: workbook is calculated, saved, Excel quit

The server is single-threaded and synchronous. Excel itself serializes
AppleScript calls, so adding threading would just add lock contention.
"""
from __future__ import annotations

import atexit
import base64
import json
import logging
import os
import queue
import shutil
import signal
import tempfile
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any

import xlwings as xw

from recalc import recalc_xlsx

log = logging.getLogger("pseudo_bridge")


# ---- xlwings backend -------------------------------------------------------


class XlwingsBackend:
    """Wraps a single hidden Excel instance + one active workbook."""

    # xlwings / Excel constants
    HALIGN = {
        "Left": -4131,
        "Center": -4108,
        "Right": -4152,
        "CenterAcrossSelection": 7,
    }
    # Border index
    BORDER_INDEX = {
        "top": 8,     # xlEdgeTop
        "bottom": 9,  # xlEdgeBottom
        "left": 7,    # xlEdgeLeft
        "right": 10,  # xlEdgeRight
    }
    XLCONTINUOUS = 1
    XLTHIN = 2

    def __init__(self, output_xlsx: Path):
        """Always start from a blank workbook.

        Excel on macOS is picky about save paths — saving into arbitrary
        user directories (under ~/Documents/<project>/...) can hang
        silently on a sandbox permission dialog. We work around this by
        giving xlwings a /tmp path to operate on, then copying to the
        caller-requested output path on shutdown and on every snapshot.
        """
        self.final_output = Path(output_xlsx)
        self.final_output.parent.mkdir(parents=True, exist_ok=True)
        # Excel-facing working path must live at a *flat* path directly under
        # /private/tmp. Hidden Excel on macOS silently refuses saves into:
        #   - paths under the user's Documents/ tree (sandbox block)
        #   - paths under /var/folders/... (sandbox block)
        #   - subdirectories of /tmp (only a `~$…xlsx` lock lands, xlsx never commits)
        # Flat /tmp/*.xlsx and /private/tmp/*.xlsx both work. We use the latter.
        self._work_dir = Path("/private/tmp")
        self.output_xlsx = self._work_dir / f"pseudo_bridge_{os.getpid()}_{uuid.uuid4().hex[:8]}.xlsx"
        self._closed = False
        self.excel_pid: int | None = None

        log.info("starting hidden Excel instance (xlwings)…")
        self.app = xw.App(visible=False, add_book=False)
        self.excel_pid = self.app.pid
        log.info(f"Excel PID={self.excel_pid}")
        try:
            self.app.display_alerts = False
        except Exception:
            pass
        try:
            self.app.api.iteration = True
        except Exception:
            pass

        log.info(f"creating empty workbook at {self.output_xlsx}")
        self.wb = self.app.books.add()
        self.wb.save(str(self.output_xlsx))
        log.info(f"backend ready; output staged at {self.output_xlsx}, will copy to {self.final_output}")

    # ---- lifecycle ----
    def _copy_to_final(self) -> None:
        """Copy the /tmp working xlsx to the caller-requested output path."""
        try:
            if self.output_xlsx.exists():
                shutil.copyfile(self.output_xlsx, self.final_output)
        except Exception as e:
            log.warning(f"copy to final failed: {e}")

    @staticmethod
    def _pid_alive(pid: int) -> bool:
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True

    def _force_kill_excel(self) -> None:
        """Kill the Excel instance by PID if app.quit() didn't take.

        Only kills the specific PID we started — other Excel instances
        (user's real work, other pseudo-bridge runs in parallel) are
        untouched.
        """
        pid = self.excel_pid
        if pid is None:
            return
        # Give quit() a grace window
        for _ in range(20):
            if not self._pid_alive(pid):
                return
            time.sleep(0.1)
        log.warning(f"Excel PID {pid} still alive after app.quit(); sending SIGTERM")
        try:
            os.kill(pid, signal.SIGTERM)
        except Exception:
            pass
        for _ in range(20):
            if not self._pid_alive(pid):
                return
            time.sleep(0.1)
        log.warning(f"Excel PID {pid} still alive after SIGTERM; sending SIGKILL")
        try:
            os.kill(pid, signal.SIGKILL)
        except Exception:
            pass

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            self.app.calculate()
            self.wb.save(str(self.output_xlsx))
            self._copy_to_final()
        except Exception as e:
            log.warning(f"final save failed: {e}")
        try:
            self.wb.close()
        except Exception:
            pass
        try:
            self.app.quit()
        except Exception:
            pass
        self._force_kill_excel()
        try:
            if self.output_xlsx.exists():
                self.output_xlsx.unlink()
            for sidecar in self._work_dir.glob(f"~${self.output_xlsx.stem}*"):
                sidecar.unlink()
        except Exception:
            pass

    # ---- helpers ----
    def _sheet(self, name: str):
        if name not in [s.name for s in self.wb.sheets]:
            raise ValueError(f"sheet '{name}' not found (have: {[s.name for s in self.wb.sheets]})")
        return self.wb.sheets[name]

    @staticmethod
    def _rgb_tuple(color: str | None) -> tuple[int, int, int] | None:
        if not color:
            return None
        c = color.lstrip("#")
        if len(c) == 6:
            return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))
        return None

    # ---- commands ----
    def createSheet(self, name: str) -> dict:
        # Idempotent: delete then create (matches live bridge semantics).
        # xlwings/macOS can fail if we delete the last remaining sheet, so
        # add a placeholder first when needed, then rename.
        existing = [s.name for s in self.wb.sheets]
        if name in existing and len(existing) == 1:
            # Can't delete the only sheet directly — add a temp, delete the
            # named one, then add the real one, then remove the temp.
            tmp = self.wb.sheets.add()
            tmp_name = tmp.name
            self.wb.sheets[name].delete()
            new_sheet = self.wb.sheets.add()
            new_sheet.name = name
            if tmp_name in [s.name for s in self.wb.sheets]:
                self.wb.sheets[tmp_name].delete()
        else:
            if name in existing:
                self.wb.sheets[name].delete()
            # Add without name, then rename (more reliable on macOS appscript).
            new_sheet = self.wb.sheets.add()
            new_sheet.name = name
        return {"ok": True, "name": name}

    def deleteSheet(self, name: str) -> dict:
        if name in [s.name for s in self.wb.sheets]:
            self.wb.sheets[name].delete()
        return {"ok": True}

    def getSheetNames(self) -> dict:
        return {"ok": True, "sheets": [s.name for s in self.wb.sheets]}

    def writeValues(self, sheet: str, address: str, values: list) -> dict:
        ws = self._sheet(sheet)
        ws.range(address).value = values
        return {"ok": True}

    def writeFormulas(self, sheet: str, address: str, formulas: list) -> dict:
        ws = self._sheet(sheet)
        # xlwings accepts a 2D list assigned to .formula
        ws.range(address).formula = formulas
        return {"ok": True}

    def readValues(self, sheet: str, address: str) -> dict:
        ws = self._sheet(sheet)
        return {"ok": True, "values": ws.range(address).value}

    def readFormat(self, sheet: str, address: str) -> dict:
        ws = self._sheet(sheet)
        rng = ws.range(address)
        fmt = {"numberFormat": rng.number_format}
        return {"ok": True, "format": fmt}

    def formatRange(self, sheet: str, address: str, format: dict) -> dict:
        ws = self._sheet(sheet)
        # address may be comma-joined (used by the live bridge's batch format grouping).
        # Apply format to each sub-range.
        subs = [a.strip() for a in address.split(",")]
        for sub in subs:
            rng = ws.range(sub)
            font = format.get("font") or {}
            if "name" in font:
                rng.font.name = font["name"]
            if "size" in font:
                rng.font.size = float(font["size"])
            if "bold" in font:
                rng.font.bold = bool(font["bold"])
            if "italic" in font:
                rng.font.italic = bool(font["italic"])
            rgb = self._rgb_tuple(font.get("color"))
            if rgb is not None:
                rng.font.color = rgb
            fill = format.get("fill") or {}
            fill_rgb = self._rgb_tuple(fill.get("color"))
            if fill_rgb is not None:
                rng.color = fill_rgb
            halign = format.get("horizontalAlignment")
            if halign and halign in self.HALIGN:
                try:
                    rng.api.HorizontalAlignment = self.HALIGN[halign]
                except Exception as e:
                    log.debug(f"halign set failed on {sub}: {e}")
            # Borders: top/bottom/left/right applied to OUTER edges of the range
            for side, idx in self.BORDER_INDEX.items():
                spec = format.get(side) if isinstance(format.get(side), (dict, bool)) else None
                if spec:
                    try:
                        bd = rng.api.Borders.Item(idx)
                        bd.LineStyle = self.XLCONTINUOUS
                        bd.Weight = self.XLTHIN
                        if isinstance(spec, dict):
                            c = self._rgb_tuple(spec.get("color"))
                            if c:
                                bd.Color = c[2] * 256 * 256 + c[1] * 256 + c[0]
                    except Exception as e:
                        log.debug(f"border {side} set failed on {sub}: {e}")
        return {"ok": True}

    def setNumberFormat(self, sheet: str, address: str, format: str) -> dict:
        ws = self._sheet(sheet)
        for sub in [a.strip() for a in address.split(",")]:
            ws.range(sub).number_format = format
        return {"ok": True}

    def setColumnWidths(self, sheet: str, columns: dict) -> dict:
        ws = self._sheet(sheet)
        for col, w in columns.items():
            try:
                ws.range(f"{col}1").column_width = float(w)
            except Exception as e:
                log.debug(f"col width {col}={w} failed: {e}")
        return {"ok": True}

    def setRowHeights(self, sheet: str, rows: dict) -> dict:
        ws = self._sheet(sheet)
        for r, h in rows.items():
            try:
                ws.range(f"A{int(r)}").row_height = float(h)
            except Exception as e:
                log.debug(f"row height {r}={h} failed: {e}")
        return {"ok": True}

    def autoFitColumns(self, sheet: str, address: str | None = None) -> dict:
        ws = self._sheet(sheet)
        if address:
            ws.range(address).autofit(axis="c")
        else:
            ws.used_range.autofit(axis="c")
        return {"ok": True}

    def autoFitRows(self, sheet: str, address: str | None = None) -> dict:
        ws = self._sheet(sheet)
        if address:
            ws.range(address).autofit(axis="r")
        else:
            ws.used_range.autofit(axis="r")
        return {"ok": True}

    def mergeCells(self, sheet: str, address: str) -> dict:
        ws = self._sheet(sheet)
        ws.range(address).merge()
        return {"ok": True}

    def freezeRows(self, sheet: str, count: int) -> dict:
        # ActiveWindow requires the app to be foregrounded. ws.activate()
        # briefly steals focus on macOS *before* raising, which is disruptive
        # when the user is on their desktop. Skip entirely for hidden apps.
        if not getattr(self.app, "visible", True):
            return {"ok": True, "skipped": "hidden_app"}
        ws = self._sheet(sheet)
        try:
            ws.activate()
            self.app.api.ActiveWindow.SplitRow = int(count)
            self.app.api.ActiveWindow.FreezePanes = True
            return {"ok": True}
        except Exception as e:
            log.debug(f"freezeRows skipped: {e}")
            return {"ok": True, "skipped": str(e)}

    def setShowGridLines(self, sheet: str, show: bool) -> dict:
        # DisplayGridlines is a window-level property — same caveat as
        # freezeRows. Skip for hidden apps to avoid macOS focus flashes.
        if not getattr(self.app, "visible", True):
            return {"ok": True, "skipped": "hidden_app"}
        ws = self._sheet(sheet)
        try:
            ws.activate()
            self.app.api.ActiveWindow.DisplayGridlines = bool(show)
            return {"ok": True}
        except Exception as e:
            log.debug(f"setShowGridLines skipped: {e}")
            return {"ok": True, "skipped": str(e)}

    def clearRange(self, sheet: str, address: str) -> dict:
        ws = self._sheet(sheet)
        ws.range(address).clear_contents()
        return {"ok": True}

    def setIterativeCalculation(self, enabled: bool, maxIteration: int = 100, maxChange: float = 0.001) -> dict:
        try:
            self.app.api.iteration = bool(enabled)
            self.app.api.max_iteration = int(maxIteration)
            self.app.api.max_change = float(maxChange)
        except Exception as e:
            log.debug(f"setIterativeCalculation failed: {e}")
        return {"ok": True}

    def batch(self, commands: list) -> dict:
        results = []
        for item in commands:
            cmd = item.get("command")
            params = item.get("params") or {}
            handler = getattr(self, cmd, None)
            if handler is None:
                results.append({"ok": False, "error": f"unknown command: {cmd}"})
                continue
            try:
                results.append(handler(**params))
            except Exception as e:
                results.append({"ok": False, "error": f"{type(e).__name__}: {e}"})
        return {"ok": True, "results": results}

    def dumpSheet(self, sheet: str) -> dict:
        ws = self._sheet(sheet)
        used = ws.used_range
        values = used.value if used.count else None
        return {"ok": True, "values": values, "address": used.address if used.count else None}

    def protectWorkbook(self, password: str | None = None) -> dict:
        # No-op for headless — the model isn't being edited by a human.
        return {"ok": True}

    def unprotectWorkbook(self, password: str | None = None) -> dict:
        return {"ok": True}

    def setStatus(self, text: str) -> dict:
        log.info(f"[builder status] {text}")
        return {"ok": True}

    def saveSnapshot(self) -> dict:
        self.app.calculate()
        self.wb.save(str(self.output_xlsx))
        self._copy_to_final()
        b64 = base64.b64encode(self.output_xlsx.read_bytes()).decode("ascii")
        return {"ok": True, "base64": b64, "path": str(self.final_output)}

    def createChart(self, **kwargs) -> dict:
        # Minimal: not implemented yet for the pseudo-bridge.
        return {"ok": False, "error": "createChart not implemented in pseudo-bridge"}

    def createTable(self, **kwargs) -> dict:
        return {"ok": False, "error": "createTable not implemented in pseudo-bridge"}


# ---- HTTP server -----------------------------------------------------------


class _Job:
    """Pending xlwings work to be executed on the worker thread."""
    __slots__ = ("kind", "payload", "done", "result", "error")

    def __init__(self, kind: str, payload: Any):
        self.kind = kind          # "command" | "checkpoint" | "shutdown"
        self.payload = payload
        self.done = threading.Event()
        self.result: Any = None
        self.error: str | None = None


class _Handler(BaseHTTPRequestHandler):
    server_version = "PseudoBridge/0.1"

    def log_message(self, fmt: str, *args: Any) -> None:  # quieter
        log.debug("http %s", fmt % args)

    def _read_json(self) -> dict:
        n = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(n) if n else b""
        try:
            return json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            return {}

    def _send_json(self, payload: dict, status: int = 200) -> None:
        # xlwings occasionally returns datetime values from Excel cells
        # (dates, serial-formatted timestamps). Fall back to str() for any
        # type the default encoder can't handle — the Builder reads values
        # as opaque data, so stringification is fine.
        data = json.dumps(payload, default=str).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _enqueue(self, kind: str, payload: Any, timeout: float = 600.0) -> dict:
        job = _Job(kind, payload)
        self.server.cmd_queue.put(job)  # type: ignore[attr-defined]
        if not job.done.wait(timeout=timeout):
            return {"ok": False, "error": "timeout waiting for xlwings worker"}
        if job.error:
            return {"ok": False, "error": job.error}
        return job.result

    def do_POST(self) -> None:
        body = self._read_json()
        path = self.path.split("?", 1)[0]
        try:
            if path == "/api/command":
                cmd = body.get("command")
                params = body.get("params") or {}
                if not cmd:
                    self._send_json({"ok": False, "error": "missing command"}, 400)
                    return
                result = self._enqueue("command", (cmd, params))
                status = 200 if result.get("ok", True) else 500
                self._send_json(result, status)
            elif path == "/api/checkpoint":
                description = body.get("description", "")
                self._enqueue("checkpoint", description)
                # Return immediately-good — matches live bridge "receipt, not verdict"
                self.server.checkpoint_log.append({"t": time.time(), "description": description})  # type: ignore[attr-defined]
                self._send_json({"ok": True, "status": "pass", "findings": []})
            elif path == "/api/emit":
                text = body.get("text", "")
                self.server.emit_log.append({"t": time.time(), "text": text})  # type: ignore[attr-defined]
                log.info(f"[builder emit] {text}")
                self._send_json({"ok": True})
            else:
                self._send_json({"ok": False, "error": f"unknown path: {path}"}, 404)
        except Exception as e:
            log.exception("handler error")
            self._send_json({"ok": False, "error": f"{type(e).__name__}: {e}"}, 500)


class PseudoBridgeServer:
    """Context-manager server.

    Architecture:
      - HTTP server runs in a background thread; accepts any number of connections.
      - A dedicated 'worker' thread owns the xlwings App + workbook. It drains
        a FIFO job queue and executes commands there. xlwings on macOS uses
        appscript under the hood, which is NOT thread-safe — so all Excel
        interactions must happen on one thread.

    Entering the context:
      1. Spawns the worker thread, which creates the App + workbook and signals ready.
      2. Spawns the HTTP thread.

    Exiting:
      1. Sends a shutdown job to the worker — it saves, closes, quits.
      2. Shuts the HTTP server.
    """

    def __init__(
        self,
        output_xlsx: Path | str,
        port: int = 3100,
        host: str = "127.0.0.1",
    ):
        self.output_xlsx = Path(output_xlsx)
        self.port = port
        self.host = host
        self.cmd_queue: queue.Queue[_Job] = queue.Queue()
        self._httpd: HTTPServer | None = None
        self._http_thread: threading.Thread | None = None
        self._worker_thread: threading.Thread | None = None
        self._worker_ready = threading.Event()
        self._worker_error: BaseException | None = None
        self._shutdown = threading.Event()
        self._prev_signal_handlers: dict[int, Any] = {}
        self.backend_ref: XlwingsBackend | None = None
        self.checkpoint_log: list[dict] = []
        self.emit_log: list[dict] = []

    @property
    def base_url(self) -> str:
        return f"http://{self.host}:{self.port}"

    # --- worker loop ---
    def _worker_loop(self) -> None:
        try:
            backend = XlwingsBackend(self.output_xlsx)
        except Exception as e:
            self._worker_error = e
            self._worker_ready.set()
            return
        self.backend_ref = backend
        self._worker_ready.set()
        try:
            while True:
                job = self.cmd_queue.get()
                if job.kind == "shutdown":
                    try:
                        backend.close()
                    finally:
                        job.done.set()
                    return
                try:
                    if job.kind == "command":
                        cmd, params = job.payload
                        handler = getattr(backend, cmd, None)
                        if handler is None:
                            job.error = f"unknown command: {cmd}"
                        else:
                            job.result = handler(**params)
                    elif job.kind == "checkpoint":
                        backend.app.calculate()
                        backend.wb.save(str(backend.output_xlsx))
                        backend._copy_to_final()
                        job.result = {"ok": True}
                    else:
                        job.error = f"unknown job kind: {job.kind}"
                except Exception as e:
                    log.exception("worker command failed")
                    job.error = f"{type(e).__name__}: {e}"
                finally:
                    job.done.set()
        except Exception:
            log.exception("worker loop crashed")
        finally:
            # Worker loop exit (clean or crash) — make sure Excel dies.
            try:
                backend.close()
            except Exception:
                pass

    def __enter__(self) -> "PseudoBridgeServer":
        self._worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._worker_thread.start()
        if not self._worker_ready.wait(timeout=180):
            raise RuntimeError("xlwings worker did not initialize within 180s")
        if self._worker_error is not None:
            raise self._worker_error

        self._httpd = HTTPServer((self.host, self.port), _Handler)
        setattr(self._httpd, "cmd_queue", self.cmd_queue)
        setattr(self._httpd, "checkpoint_log", self.checkpoint_log)
        setattr(self._httpd, "emit_log", self.emit_log)
        self._http_thread = threading.Thread(target=self._httpd.serve_forever, daemon=True)
        self._http_thread.start()

        # Register cleanup hooks so Excel dies even on SIGINT/SIGTERM/os._exit
        atexit.register(self._emergency_cleanup)
        for sig_num in (signal.SIGINT, signal.SIGTERM):
            try:
                prev = signal.signal(sig_num, self._on_signal)
                self._prev_signal_handlers[sig_num] = prev
            except (ValueError, OSError):
                # signal() only works on main thread — fine if this isn't main
                pass

        log.info(f"pseudo bridge listening at {self.base_url} — output → {self.output_xlsx}")
        return self

    def _on_signal(self, signum: int, frame: Any) -> None:
        log.warning(f"received signal {signum}; cleaning up pseudo-bridge")
        try:
            self._cleanup()
        finally:
            # Restore previous handler and re-raise so default behavior runs
            prev = self._prev_signal_handlers.get(signum, signal.SIG_DFL)
            try:
                signal.signal(signum, prev)  # type: ignore[arg-type]
            except Exception:
                pass
            os.kill(os.getpid(), signum)

    def _cleanup(self) -> None:
        # Drain: ask worker to close (best effort; if blocked, backend_ref.close()
        # from the backend side + force-kill will finish the job).
        try:
            shutdown_job = _Job("shutdown", None)
            self.cmd_queue.put(shutdown_job)
            shutdown_job.done.wait(timeout=10)
        except Exception:
            log.exception("shutdown job failed")
        # If the worker is wedged, call close() directly on the backend.
        # close() is idempotent and handles force-kill.
        try:
            if self.backend_ref is not None:
                self.backend_ref.close()
        except Exception:
            pass
        try:
            if self._httpd is not None:
                self._httpd.shutdown()
                self._httpd.server_close()
        except Exception:
            log.exception("httpd shutdown failed")

    def _emergency_cleanup(self) -> None:
        """atexit hook — runs if __exit__ didn't (e.g. process dying ungracefully)."""
        if self._shutdown.is_set():
            return
        self._shutdown.set()
        try:
            self._cleanup()
        except Exception:
            pass

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._shutdown.is_set():
            return
        self._shutdown.set()
        # Unregister atexit so we don't double-cleanup on interpreter exit
        try:
            atexit.unregister(self._emergency_cleanup)
        except Exception:
            pass
        # Restore signal handlers
        for sig_num, prev in self._prev_signal_handlers.items():
            try:
                signal.signal(sig_num, prev)
            except Exception:
                pass
        self._prev_signal_handlers.clear()
        self._cleanup()


# ---- CLI ------------------------------------------------------------------


def _main() -> None:
    import argparse
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", required=True, type=Path)
    ap.add_argument("--port", type=int, default=3100)
    args = ap.parse_args()
    with PseudoBridgeServer(output_xlsx=args.output, port=args.port) as s:
        log.info(f"server up — BRIDGE_URL=http://127.0.0.1:{args.port}")
        log.info("press Ctrl+C to stop")
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("stopping…")


if __name__ == "__main__":
    _main()
