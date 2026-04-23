# User context

You are an analyst at a lender building a standard amortization schedule for
a $1M term loan at 6.0% annual rate over 60 monthly periods.

## Defaults you'd pick

- Equal monthly payments via Excel's `PMT(rate/12, 60, -principal)` — return
  is positive.
- End-of-period interest: interest_t = beginning_balance_t × (annual_rate/12).
- Simple monthly compounding — no 30/360 day-count adjustment needed.
- Follow `inputs/schedule_format.md` exactly for columns and formatting.
- Final balance should hit zero within rounding (within $0.01 is acceptable).

## Things you do NOT care about

- Cell colors beyond the header row fill.
- Exact column widths (ballpark per the format doc is fine).
- Whether the parameter block lives above or beside the table, so long as
  formulas reference it and the sheet is named exactly `Amortization Schedule`.
