# Task: build a Returns Analysis sheet

Build a `Returns Analysis` sheet showing each cap table holder's payout at
5 exit valuations: $100M, $250M, $500M, $1B, $2B.

- Read `inputs/cap_table.xlsx` for holder shares.
- Follow `inputs/style_guide.md` for layout and formatting.
- Use formulas (no literals) for % FD, payouts, and totals. Shares should
  reference the Cap Table (import the sheet into the workbook or cross-
  reference it).
- Write `results.json` mapping each output key to its `Sheet!Cell`.

Keep the sheet named exactly `Returns Analysis`.
