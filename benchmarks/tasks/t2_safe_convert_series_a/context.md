# User context

You are a VC associate modeling Rivet Labs' Series A. The company has two
SAFEs on the balance sheet; you want to show the founder what happens at
close.

## Defaults you'd pick

- Pre-round FD = the Total FD row of the input cap table = 10,500,000 shares.
  This is the base for all SAFE and headline calculations.
- Headline price uses the simple formula: pre_money / pre_round_FD.
- SAFE conversion: for each SAFE, compute cap_price and discount_price, take
  the lower as conversion price, then shares = investment / conversion price.
  (Exact formulas are in `safe_terms.md`.)
- Option pool top-up: 10% post-money target, pre-money inclusion, **net** of
  existing unallocated (500K counts toward the target).
- Post-money FD = pre-round FD + SAFE shares + new investor shares + pool top-up.

## Things you do NOT care about

- Color/formatting beyond clear section headers.
- Exact row numbers — just want the labels clearly there.
- Whether percentages are shown with 1 or 2 decimals.
- Whether you use named ranges.

## Things you DO care about

- Every derived number must be a formula. If I change pre-money in Series A
  cell B3, the whole model should recompute — no hardcoded literals in
  downstream cells.
- Cross-sheet refs: SAFE Convert reads pre-round FD from Cap Table and
  headline price from Series A. Series A reads total SAFE shares from SAFE
  Convert. Do not duplicate these numbers as literals on multiple sheets.
