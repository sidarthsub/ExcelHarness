# Task: option pool top-up new-share issuance

The `Calc` sheet in `stubs/starter.xlsx` has the inputs already laid out:

- `B2` — existing fully diluted shares before the round
- `B3` — existing unallocated option pool shares
- `B4` — new investor shares being issued in this round (already determined)
- `B5` — target option pool size as a % of **post-money** fully diluted

Compute the number of **new** option-pool shares the company must issue so that
the total option pool (existing unallocated + new issuance) equals the target
% of the post-money fully diluted share count.

Put the answer in **cell `B10`** of the `Calc` sheet. Use a formula that
references the inputs — do not hardcode a literal.

Reference: `inputs/pre_cap_table.xlsx` shows the current cap table for context.
