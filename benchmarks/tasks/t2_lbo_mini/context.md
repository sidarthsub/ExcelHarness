# User context

PE associate building a standard LBO for Reed Industries.

## Defaults you'd pick

- 100% cash sweep on TLB after mandatory amort.
- Mezz doesn't amortize; constant $50M balance through exit.
- Exit proceeds = Y5 EBITDA × exit multiple − net debt at exit.
- Formulas only for derived values — no hardcoded downstream numbers.
- Sponsor equity is the plug in the Sources & Uses sheet.

## Things you do NOT care about

- Color / formatting beyond clear section headers.
- Exact row numbers — labels matter, positions don't.
- Named ranges vs. cell refs.

## Things you DO care about

- Every derived number is a formula, not a literal.
- TLB balance rolls forward: Y2 beginning TLB = Y1 ending TLB, etc.
- Cash sweep floored at 0 and capped so TLB can't go negative:
  `sweep = MAX(0, MIN(FCF − mandatory, TLB_beg − mandatory))`.
- Interest on TLB uses beginning-of-period balance (not average).
