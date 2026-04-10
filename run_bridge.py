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
