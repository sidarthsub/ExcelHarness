# User context

You are an associate completing a DCF. You want the standard Gordon-growth
terminal value at the end of the explicit forecast period and a clean
present-value pull-back using the horizon-year discount factor.

## Defaults you would pick

- Gordon growth formula: `TV = FCF x (1+g) / (WACC - g)` using Y5 FCF as base.
- End-of-period discounting: `factor = 1 / (1+WACC)^horizon`.
- No rounding — keep full precision.
- Formula-driven cells only (no hardcoded literals).

## Things you do NOT know or care about

- Cell formatting, number precision beyond a handful of decimals.
- Whether formulas use direct cell refs or named ranges — either is fine.
- Whether intermediate cells show up alongside the answer — you just want
  the three answers in `B9:B11`.
