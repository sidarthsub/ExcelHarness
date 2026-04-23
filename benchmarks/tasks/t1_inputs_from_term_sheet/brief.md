# Task: build a Series A Inputs sheet from the term sheet

Read `inputs/term_sheet.md` and build an `Inputs` sheet that captures every
Series A parameter a downstream Series A model would need.

Organize into clearly labeled sections:

1. Round sizing (pre-money, round size, post-money)
2. Investor allocations (each investor with amount, % of round, and a total)
3. Cap table mechanics (pre-round FD, pre-round unallocated, target pool %)
4. Preference terms (liquidation preference, dividend rate)

Put the file at `output/model.xlsx`. All derived values (post-money,
percentages, totals) must be **formulas**, not literals — those are the
cells a downstream model will reference.

Keep the sheet named exactly `Inputs`.
