"""Phase 1 smoke test: hand-written script that exercises the bridge against live Excel.

Before running:
  1. Start the bridge server: python3 run_bridge.py
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
