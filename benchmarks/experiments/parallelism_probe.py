"""Probe: can we run N PseudoBridgeServer instances concurrently on this machine?

Each server spawns its own hidden Excel process via `xw.App(visible=False)`.
On macOS, whether 2+ such processes coexist peacefully depends on the
installed Excel build and AppleScript permissions. This script answers
the question empirically without spending any LLM dollars.

The probe:

  1. Start N PseudoBridgeServer instances on distinct ports, each
     pointing at a distinct output xlsx.
  2. Confirm their reported Excel PIDs are all distinct.
  3. From N worker threads in parallel, drive each server with a
     uniquely-identifiable write: server i writes `"probe_i"` to
     `Sheet1!A1`. We use the real `bridge.Bridge` HTTP client so the
     exercise exactly matches how a Builder would talk to it.
  4. Shut down each server (this forces a save + Excel.quit()).
  5. Load every output xlsx with openpyxl and assert:
       - output_i!Sheet1!A1 == "probe_i"      (no lost writes)
       - no other probe_j marker leaked in    (no crosstalk)

Exit code 0 if all N instances behaved independently; 1 otherwise.

Default N=2. Bump via `--n 3` etc. Wall time is dominated by Excel
cold-start (~5-10s per instance on macOS).
"""
from __future__ import annotations

import argparse
import logging
import shutil
import socket
import sys
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import openpyxl

# Make sure the repo root is on sys.path before importing bridge/pseudo_bridge.
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from bridge import Bridge  # noqa: E402
from benchmarks.pseudo_bridge import PseudoBridgeServer  # noqa: E402

log = logging.getLogger("parallelism_probe")


def _pick_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _run_one_worker(idx: int, base_url: str) -> dict:
    """Drive a single server with a unique marker."""
    t0 = time.time()
    marker = f"probe_{idx}"
    try:
        b = Bridge(base_url=base_url, verify_tls=False, timeout=60.0)
        # A server always starts with a blank workbook named Sheet1.
        # We rename it to make the test content-visible, then write.
        # Using write_values is the minimal bridge primitive that exercises
        # an AppleScript → Excel round-trip.
        b.write_values("Sheet1", "A1", [[marker]])
        b.write_values("Sheet1", "B1", [[idx]])
        # Force a save-to-caller-path so openpyxl can read it after shutdown.
        b.checkpoint(f"probe idx={idx}")
        return {"idx": idx, "marker": marker, "ok": True,
                "wall_s": round(time.time() - t0, 2)}
    except Exception as e:
        return {"idx": idx, "marker": marker, "ok": False,
                "error": f"{type(e).__name__}: {e}",
                "wall_s": round(time.time() - t0, 2)}


def probe(n: int, workdir: Path) -> dict:
    servers: list[PseudoBridgeServer] = []
    output_paths: list[Path] = []
    ports: list[int] = []
    excel_pids: list[int | None] = []

    # --- Phase 1: start servers sequentially ---
    # We intentionally start them one at a time (the outer loop runs serial)
    # since the concurrency we're testing is *command servicing*, not
    # *server startup*. If Excel can't even start a second hidden instance,
    # we fail here and the error is diagnostic.
    log.info(f"starting {n} PseudoBridgeServer instances…")
    start_t0 = time.time()
    start_errors = []
    for i in range(n):
        port = _pick_free_port()
        out = workdir / f"probe_out_{i}.xlsx"
        try:
            s = PseudoBridgeServer(output_xlsx=out, port=port, host="127.0.0.1")
            s.__enter__()
        except Exception as e:
            start_errors.append({"idx": i, "error": f"{type(e).__name__}: {e}"})
            log.error(f"startup failed for instance {i}: {e}")
            break
        servers.append(s)
        output_paths.append(out)
        ports.append(port)
        backend = getattr(s, "backend_ref", None)
        pid = backend.excel_pid if backend else None
        excel_pids.append(pid)
        log.info(f"  instance {i}: port={port} excel_pid={pid} out={out}")

    startup_wall = round(time.time() - start_t0, 2)

    if start_errors or len(servers) < n:
        # Tear down whatever came up.
        for s in servers:
            try: s.__exit__(None, None, None)
            except Exception: pass
        return {
            "ok": False,
            "phase": "startup",
            "started": len(servers),
            "requested": n,
            "startup_wall_seconds": startup_wall,
            "startup_errors": start_errors,
            "excel_pids": excel_pids,
        }

    distinct_pids = len({p for p in excel_pids if p is not None}) == n
    log.info(f"distinct_excel_pids={distinct_pids} pids={excel_pids}")

    # --- Phase 2: drive servers in parallel via threads ---
    log.info(f"firing {n} concurrent write+checkpoint workers…")
    work_t0 = time.time()
    worker_results: list[dict] = []
    with ThreadPoolExecutor(max_workers=n) as pool:
        futs = [pool.submit(_run_one_worker, i, servers[i].base_url) for i in range(n)]
        for fut in as_completed(futs):
            worker_results.append(fut.result())
    worker_results.sort(key=lambda r: r["idx"])
    work_wall = round(time.time() - work_t0, 2)

    all_workers_ok = all(r["ok"] for r in worker_results)
    log.info(f"workers: all_ok={all_workers_ok} wall={work_wall}s")

    # --- Phase 3: shut down (this flushes the xlsx to disk) ---
    log.info("shutting down servers…")
    shutdown_t0 = time.time()
    for s in servers:
        try:
            s.__exit__(None, None, None)
        except Exception as e:
            log.warning(f"shutdown error: {e}")
    shutdown_wall = round(time.time() - shutdown_t0, 2)

    # --- Phase 4: verify outputs ---
    verification = []
    for i, out in enumerate(output_paths):
        row = {"idx": i, "path": str(out), "exists": out.exists()}
        if not out.exists():
            row["ok"] = False
            row["reason"] = "output file missing"
            verification.append(row)
            continue
        try:
            wb = openpyxl.load_workbook(out, data_only=True)
            ws = wb[wb.sheetnames[0]]
            own_marker = ws["A1"].value
            own_idx = ws["B1"].value
            # Scan all sheets for leaked markers from other instances.
            leaked = []
            for sheet_name in wb.sheetnames:
                ws2 = wb[sheet_name]
                for rrow in ws2.iter_rows(values_only=True):
                    for v in rrow:
                        if isinstance(v, str) and v.startswith("probe_") and v != f"probe_{i}":
                            leaked.append(v)
            wb.close()
            row["own_marker"] = own_marker
            row["own_idx_cell"] = own_idx
            row["leaked_markers"] = leaked
            row["ok"] = (own_marker == f"probe_{i}") and (own_idx == i) and not leaked
            if not row["ok"]:
                row["reason"] = (
                    f"own_marker={own_marker!r} expected {f'probe_{i}'!r}; "
                    f"leaked={leaked}"
                )
        except Exception as e:
            row["ok"] = False
            row["reason"] = f"{type(e).__name__}: {e}"
        verification.append(row)

    all_verified = all(r["ok"] for r in verification)

    return {
        "ok": distinct_pids and all_workers_ok and all_verified,
        "n": n,
        "startup_wall_seconds": startup_wall,
        "worker_wall_seconds": work_wall,
        "shutdown_wall_seconds": shutdown_wall,
        "excel_pids": excel_pids,
        "distinct_excel_pids": distinct_pids,
        "workers": worker_results,
        "verification": verification,
    }


def _main() -> int:
    import json
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    ap = argparse.ArgumentParser(prog="benchmarks.experiments.parallelism_probe")
    ap.add_argument("--n", type=int, default=2,
                    help="How many PseudoBridgeServer instances to spin up in parallel.")
    ap.add_argument("--keep-outputs", action="store_true",
                    help="Keep the probe's output xlsx files for inspection.")
    args = ap.parse_args()

    with tempfile.TemporaryDirectory(prefix="parallelism_probe_") as td:
        workdir = Path(td)
        report = probe(args.n, workdir)
        print(json.dumps(report, indent=2, default=str))
        if args.keep_outputs:
            dest = Path("/tmp") / f"parallelism_probe_{int(time.time())}"
            shutil.copytree(workdir, dest)
            print(f"\noutputs kept at {dest}", file=sys.stderr)
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(_main())
