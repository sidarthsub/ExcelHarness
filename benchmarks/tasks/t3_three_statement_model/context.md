# User context

You are an associate at a growth PE fund. You're building a quick
3-statement model for Bluebird SaaS to evaluate a potential investment.
You want it clean and auditable.

## Defaults you would pick

- 4 columns: 2026 (historical) + 2027, 2028, 2029 (forecast).
- Every statement line comes from `Inputs` via formula — no literals in the
  statement sheets.
- Tax is on positive pre-tax only; no NOL carryforward tracked.
- Debt is flat — no amortization, no new draws. Interest = 6% × opening debt.
- A/R = Revenue × DSO/365; A/P = COGS × DPO/365 (simple working capital).
- No inventory line needed — it's SaaS.
- PP&E net roll: prior + CapEx − D&A. Straight-line D&A, flat $1M/year.
- Retained earnings in 2026 is the plug that makes the historical BS
  balance. In forecast years, RE rolls with net income.
- Ending cash on the CF ties to cash on the BS every period.

## Things you do NOT care about

- Color, fonts, borders beyond clear visual separation of totals.
- Whether you use negative signs or parentheses for negatives.
- Whether the CF uses direct or indirect method — indirect is fine.
- NOL carry-forwards, deferred taxes, stock-based comp, restructuring
  charges — keep it clean.

## Things you DO care about

- BS balances every period.
- Ending cash tie between CF and BS.
- Formula-driven — changing a driver on Inputs should ripple everywhere.
