# Returns table — style conventions

## Layout
- Title row at top: "Returns Analysis" (bold, size 13)
- Header row with "Holder", "Shares", "% FD", then one column per exit
  valuation: "$100M", "$250M", "$500M", "$1B", "$2B"
- One row per holder from the cap table
- Final row: totals

## Formats
- Share counts: `#,##0`
- % FD: `0.00%`
- Dollar payouts: `$#,##0`
- Header row: bold, light-blue fill (`#DDEEFF`)
- Totals row: bold, top border

## Values
- % FD = shares / total_FD (formula referencing total)
- Payout at valuation V = shares × V / total_FD (pro-rata to FD, no
  waterfall preference in this task)
