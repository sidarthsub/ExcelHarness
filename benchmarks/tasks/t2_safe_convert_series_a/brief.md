# Task: SAFE conversion into a Series A

Build a three-sheet model showing how two outstanding SAFEs convert at a
Series A close. Save to `output/model.xlsx`.

## Sheets

1. **Cap Table** — reproduce the pre-round cap table from
   `inputs/pre_cap_table.xlsx`. Holder, class, share count, % FD, total FD.
   Match the structure of the source.

2. **SAFE Convert** — for each of the two SAFEs described in
   `inputs/safe_terms.md`:
   - inputs (investment, cap, discount)
   - cap price, headline price reference, discount price, conversion price
     (the min of cap and discount)
   - shares issued
   End with a "Total SAFE shares" cell.

3. **Series A** — round mechanics from `inputs/round_terms.md`:
   - pre-money, round size, option pool target %, pre-unallocated
   - headline price per share (derived from pre-money / pre-round FD)
   - SAFE shares (reference `SAFE Convert`)
   - new investor shares (derived)
   - option pool top-up (derived from the 10% post-money target, pre-money
     inclusion, net of existing unallocated)
   - post-money fully diluted shares

All derived values must be formulas. Cross-sheet references should be used
where indicated.
