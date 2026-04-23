# Task: build a 60-month debt amortization schedule

Build a 60-month debt amortization schedule on a sheet named
`Amortization Schedule`. Principal $1,000,000, annual rate 6.0%, 60 monthly
periods. Follow `inputs/schedule_format.md` for columns, formulas, and
formatting. Use Excel's `PMT` function for the monthly payment. Include
summary rows for total interest paid, total principal paid, and final
balance. Write `results.json` mapping each semantic key to its `Sheet!Cell`
address.

Put the file at `output/model.xlsx`. The monthly payment, per-row values,
and summary totals must all be **formulas** that reference the parameter
block — not hardcoded literals.
