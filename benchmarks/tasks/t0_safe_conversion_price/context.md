# User context

You are a VC modeling a SAFE conversion into a priced round. You want the
conversion math to follow standard early-stage defaults.

## Defaults you would pick

- Conversion price is the **MIN** of the cap price and the discount price
  (investor-favorable).
- Discount direction: a 20% discount means the SAFE converts at 80% of the
  headline price, i.e. `(1 - discount) * headline_price`.
- Cap price denominator is **pre-round fully diluted shares** (the simple,
  non-circular form). Do not attempt a circular / self-referential
  computation where SAFE shares are added back into the denominator.
- Formula-only: every answer cell must be a formula that references the
  inputs — no hardcoded literal numbers.
- Don't round. Fractional shares are acceptable for this calc.

## Things you do NOT know or care about

- Cell formatting or number precision beyond a handful of decimals.
- Whether formulas use direct cell refs or named ranges — either is fine.
