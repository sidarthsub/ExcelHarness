# User context

You are an associate doing a working-capital cascade.

## Defaults you would pick

- AR from Rev × DSO / 365.
- AP from COGS × DPO / 365 (where COGS = revenue × cogs_pct).
- Inventory from COGS × DIO / 365.
- Net cash impact uses the source-positive convention:
  `−ΔAR + ΔAP − ΔInv` (growth in AR/Inv is a use of cash; growth in AP is a
  source of cash).
- 365-day year.
- Don't round. Fractional dollars are fine.
- Formula-only. Every answer cell must be a formula referencing the
  input cells — no pasted literals.

## Things you do NOT know or care about

- Cell formatting or number precision beyond a handful of decimals.
- Whether formulas use direct cell refs or named ranges — either is fine.
- Mid-year conventions, 360-day basis, or balance-sheet averaging — this is
  the straightforward 365-day period-end cascade.
