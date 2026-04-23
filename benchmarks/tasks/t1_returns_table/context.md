# User context

You are a VC associate preparing a returns analysis. The cap table is
final and lives in `inputs/cap_table.xlsx`. You want a clean one-sheet
view of what each holder would receive at a handful of exit valuations.

## Defaults you'd pick

- Pure pro-rata on fully diluted shares — no liquidation preferences, no
  waterfall, no option strike netting. Every share (including unvested
  option pool) gets the same per-share payout.
- Formulas should reference the Cap Table rather than hardcoded literals
  for share counts. Import the Cap Table into the workbook (or cross-
  reference it) so the returns sheet updates if the cap table changes.
- Exactly 5 exit scenarios: $100M, $250M, $500M, $1B, $2B.
- Follow `inputs/style_guide.md` exactly for layout, headers, and number
  formats.

## Things you do NOT care about

- Font family or exact column widths beyond readability.
- Whether the Cap Table sheet is a copy or a live link — either is fine.
- Whether $1B is written as "$1B" or "$1,000M" in the header, as long as
  the value matches.
