# Task: SAFE conversion price at a priced round

Compute the SAFE conversion price and shares issued at the priced round.

The `Calc` sheet in `stubs/starter.xlsx` has the inputs already laid out in
`B3:B7`:

- `B3` — SAFE investment ($)
- `B4` — SAFE valuation cap ($, pre-money)
- `B5` — SAFE discount (decimal, e.g. 0.20 for 20%)
- `B6` — Pre-round fully diluted shares
- `B7` — Round pre-money valuation ($)

Fill `B10:B14` with formulas (no literals) that reference the inputs:

- `B10` — Headline price = pre_money / pre_FD
- `B11` — Cap price = cap / pre_FD
- `B12` — Discount price = (1 - discount) * headline_price
- `B13` — Conversion price = MIN(cap_price, discount_price)
- `B14` — Shares issued = investment / conversion_price

Write `results.json` alongside the workbook mapping the five output keys
(`headline_price`, `cap_price`, `discount_price`, `conversion_price`,
`shares_issued`) to their corresponding `Sheet!Cell` addresses.
