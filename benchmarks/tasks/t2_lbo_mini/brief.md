# Task: Mini LBO model — Reed Industries

Build a 3-sheet LBO model for Reed Industries: "Sources and Uses", "Operating
Model" (5-year projection with TLB amortization + cash sweep), and "Returns".
Follow `inputs/deal_terms.md` for deal parameters + debt mechanics, and
`inputs/model_format.md` for layout.

Use formulas — sponsor equity is a plug, TLB sweeps excess FCF after mandatory
amort. Mezzanine does not amortize.

Save to `output/model.xlsx` and write `output/results.json` mapping all 12
output keys to their Sheet!Cell addresses.
