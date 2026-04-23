# User context

You are a VC associate modeling a two-round progression (Series A then B).

## Defaults you'd pick

- Series A: simple pricing (`pre_money / pre_FD`), no pool change.
- Series B: pricing uses post-A FD as the denominator (non-circular).
- Series B pool top-up: 10% post-money target, pre-money inclusion, net of
  pre-round unallocated pool.
- Cross-sheet refs required: Series A reads Cap Table; Series B reads
  Series A. No hardcoded totals — reference the cells.
- Formulas only for derived values.
