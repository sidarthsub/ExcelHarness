# Task: YoY working capital cascade

Compute AR/AP/Inventory each year, their YoY deltas, and net WC cash impact.
Inputs in `Calc!B3:B8`. Fill `Calc!B11:B20` with **formulas** that reference
the inputs — do not hardcode literals.

Inputs:

- `B3` — Revenue Y0
- `B4` — Revenue Y1
- `B5` — COGS % of revenue
- `B6` — DSO (days)
- `B7` — DPO (days)
- `B8` — DIO (days)

Answer cells:

- `B11` — AR Y0
- `B12` — AR Y1
- `B13` — ΔAR
- `B14` — AP Y0
- `B15` — AP Y1
- `B16` — ΔAP
- `B17` — Inv Y0
- `B18` — Inv Y1
- `B19` — ΔInv
- `B20` — Net WC cash impact

Write `results.json` next to the workbook mapping each output key
(`ar_y0`, `ar_y1`, `delta_ar`, `ap_y0`, `ap_y1`, `delta_ap`, `inv_y0`,
`inv_y1`, `delta_inv`, `net_wc_cash_impact`) to the `Sheet!Cell` address
containing the answer.
