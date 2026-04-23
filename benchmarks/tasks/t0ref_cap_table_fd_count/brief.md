# Task: total fully diluted share count from reference cap table

Read the reference `inputs/messy_cap_table.xlsx`. Compute the total fully
diluted share count across all holders and classes.

Put the result in **`Calc!B5`** as a **formula** (not a literal).

Recommended: import the Cap Table sheet from the reference into the model
workbook (via `bridge.copy_sheet_from_input`) and sum its shares column.

Write `results.json` with the output key `total_fd` mapped to the cell
address holding the formula.
