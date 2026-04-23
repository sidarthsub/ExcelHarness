# Task: pro-rata allocation of a $15M follow-on pool

The `Calc` sheet in `stubs/starter.xlsx` has the inputs already laid out:

- `B3` — pool to allocate ($15M)
- `B4` — Nimbus Capital prior investment ($20M)
- `B5` — Arc Ventures prior investment ($10M)
- `B6` — Orbit Partners prior investment ($5M)
- `B7` — total prior investment (`=SUM(B4:B6)`)

Compute each investor's pro-rata share of the $15M pool, where each
allocation equals `pool × investor_prior / total_prior`.

Fill in the answers on the `Calc` sheet:

- `B10` — Nimbus allocation
- `B11` — Arc allocation
- `B12` — Orbit allocation

Use formulas that reference the input cells — do not hardcode literal
dollar amounts.

Write a `results.json` alongside your workbook mapping each output key
to its `Sheet!Cell` address:

```json
{
  "nimbus_allocation": "Calc!B10",
  "arc_allocation": "Calc!B11",
  "orbit_allocation": "Calc!B12"
}
```
