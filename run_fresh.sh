#!/bin/bash
set -euo pipefail

ROOT="/Users/sidsub/Documents/ExcelHarness"
STAGE="/tmp/excel_harness_fresh"
SRC_RUN="$ROOT/runs/20260328-194442/input"

rm -rf "$STAGE"
mkdir -p "$STAGE/input" "$STAGE/reference"

cp "$SRC_RUN/Sidekick Cap Table[50].xlsx" "$STAGE/input/"
cp "$SRC_RUN/Think More Security - SAFE (Form)[28].pdf" "$STAGE/input/"
cp "$SRC_RUN/Fig & 1011 Returns Analysis_v01.xlsx" "$STAGE/reference/"

cd "$ROOT"
python3 harness.py \
  "Make a sidekick returns analysis based on the Fig one. Sheets should exactly match the fig formatting. Copy the Sidekick Cap Table sheet directly from the input Sidekick xlsx with its existing formatting — do not rebuild it. The analysis should have a modeled out Series A (60 pre, 20M total, 1011 12, Accomplice 5, Other Series A 3, option pool to 10% post money accounting for previously unallocated), Series B (150 pre, 30 total, Other Series B 15, remaining 15 split pro rata between last round participants, option pool gross top-up to 7% post money — issue new shares equal to 7% of post-B fully diluted shares regardless of existing unallocated), post A and B scenarios. Group small common holders together, group small investors together, and group the accomplice funds together with their new investments." \
  "$STAGE/input" \
  "$STAGE/reference" \
  2>&1 | tee /tmp/harness_fresh.log
