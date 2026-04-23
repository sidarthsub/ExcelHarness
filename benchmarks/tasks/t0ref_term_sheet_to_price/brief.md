# Task: Series A share-price math from a term sheet

Read `inputs/term_sheet.md` to extract pre-money, round size, and pre-round FD
shares. Populate `B3:B5` with those values (literals) and `B8:B12` with
formulas computing post-money, share price, new investor shares, post-round
FD, and Starlight's post-round ownership. Write `results.json`.

## Stub layout (`Calc` sheet)

- `B3` — Pre-money valuation (literal, from term sheet)
- `B4` — New money / round size (literal)
- `B5` — Pre-round fully diluted shares (literal)
- `B8` — Post-money valuation (formula)
- `B9` — Share price (formula)
- `B10` — New investor shares (formula)
- `B11` — Post-round FD shares (formula)
- `B12` — Starlight % post-round (formula, decimal fraction)

All five derived cells must be formulas that reference `B3:B5` (and each other
as needed). Do not hardcode derived values.
