# Task: 1x non-participating waterfall

Compute the 1x non-participating waterfall at the given exit.

Inputs sit in `Calc!B3:B6`:

- `B3` — LP investment (original $)
- `B4` — LP post-money ownership (%)
- `B5` — Pref multiple (1.0 = 1x)
- `B6` — Exit value ($)

Fill `B10:B13` with **formulas** that reference the inputs — do not hardcode
literals:

- `B10` — preference amount = investment × multiple
- `B11` — conversion amount = ownership × exit
- `B12` — LP actual proceeds = MAX(preference, conversion)
- `B13` — common proceeds = exit − LP proceeds

Write `results.json` next to the workbook mapping each output key
(`pref_amount`, `convert_amount`, `lp_proceeds`, `common_proceeds`) to the
`Sheet!Cell` address containing the answer.
