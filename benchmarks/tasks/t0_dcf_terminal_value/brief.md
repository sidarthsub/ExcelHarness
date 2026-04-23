# Task: DCF Gordon-growth terminal value and PV

Compute the Gordon-growth terminal value at the end of Year 5, the Year 5
discount factor, and the present value of the terminal value.

Inputs are laid out on the `Calc` sheet in `B3:B6`:

- `B3` — Year 5 free cash flow
- `B4` — Terminal growth rate (g)
- `B5` — WACC
- `B6` — Horizon (years)

Fill `B9:B11` with formulas that reference the inputs — do not hardcode
literals:

- `B9`  — Terminal value (Y5)
- `B10` — Discount factor (Y5)
- `B11` — PV of terminal value

Write a `results.json` file next to your workbook mapping
`terminal_value_y5`, `discount_factor_y5`, and `pv_terminal_value` to
their `Sheet!Cell` addresses.
