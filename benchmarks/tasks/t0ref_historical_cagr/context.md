# User context

Associate reviewing historical performance.

## Defaults you would pick

- CAGR formula = `(end / start) ^ (1 / n) - 1`.
- "3-year CAGR" means 3 compounding periods (2022 → 2023 → 2024 → 2025),
  so start=2022, end=2025, n=3.
- Don't round. Keep full precision.
- Formulas reference the Revenue sheet — don't hardcode revenue values
  in the Calc sheet.

## Things you do NOT know or care about

- Cell formatting beyond what's already set.
- Whether extra helper cells exist alongside the answers — just want the
  right number at each output key.
