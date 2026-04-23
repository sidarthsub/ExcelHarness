# Task: total cash runway with a stepped burn

Compute total runway given starting cash and a two-phase burn. Phase 1 lasts
6 months at $1M/mo; phase 2 runs at $1.5M/mo indefinitely. Fill `B9`, `B10`,
`B11` on the `Calc` sheet with formulas referencing the inputs — do not
hardcode literals. Write `results.json` mapping each output key to its cell.

Outputs (on `Calc`):

- `B9`  — cash at end of phase 1
- `B10` — phase 2 months of runway at the stepped-up burn
- `B11` — total runway in months (headline answer)
