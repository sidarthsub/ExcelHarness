# Financial Model Conventions

## Structure
- One logical concept per sheet
- First row of each sheet is a header row with the sheet title in A1 or B1
- Column A: row tags/keys (short identifiers for CONCATENATE references, can be hidden). Column B: row labels. Remaining columns: data
- Inputs/assumptions should live on their own sheet, separate from calculations
- Use blank rows to separate logical sections (e.g., between Common Stock and Preferred sections)

## Formulas
- NO hardcoded numbers in formula cells — every input must trace back to a clearly labeled source cell or input sheet
- Cross-sheet references must use the full `SheetName!CellRef` syntax
- Circular references are allowed — iterative calculation is always enabled on the workbook:
  ```python
  wb.calculation.iterate = True
  wb.calculation.iterateCount = 100
  wb.calculation.iterateDelta = 0.001
  ```
- Use absolute references (`$`) for cells that shouldn't shift when copied

## Formatting — Required for Every Sheet

### Fonts
- Default: Calibri 11pt
- Sheet title: Calibri 11pt Bold (or larger if matching a reference)
- Section headers: Bold
- Totals rows: Bold

### Column Widths
- Set explicit widths for EVERY column — never leave as default
- Label columns (A/B): wide enough for the longest label (typically 30-45)
- Data columns: consistent width across all data columns (typically 12-18)
- Spacer columns: narrow (2-3) for visual separation between column groups

### Number Formats
- Currency/dollars: `#,##0` (no decimals for large figures), `#,##0.00` for per-unit
- Percentages: `0.0%` or `0.00%`
- Share counts: `#,##0`
- Ratios/multiples: `0.0x` or `#,##0.00x`
- Apply number formats to ALL numeric cells — never leave as "General"

### Cell Fills
- Input/assumption cells: light green fill (`#C6EFCE`) or light blue (`#BDD7EE`) — visually distinguish inputs from calculations
- Header rows: light gray fill (`#D9D9D9`) if used in reference
- Match the reference model's color scheme exactly when one is provided

### Borders
- Thin bottom border above subtotal rows
- Thin top + bottom border (or double bottom) on grand total rows
- Use borders to visually separate sections

### Alignment
- Text: left-aligned
- Numbers: right-aligned
- Column headers: center-aligned
- Indent sub-items with alignment indent (1-2 levels)
- NEVER merge cells. Use `centerContinuous` alignment across a range instead (`cell.alignment = Alignment(horizontal='centerContinuous')`)

### Row Heights
- Default 15pt for data rows
- Adjust for headers or title rows as needed

### Gridlines
- Sheet gridlines should be OFF (`ws.sheet_view.showGridLines = False` in openpyxl)
- Use explicit borders instead for structure

## Naming
- Sheet names: short, no spaces preferred (e.g., "Inputs", "Revenue", "OpEx")
- No special characters in sheet names
- Keep names under 20 characters

## Quality Checks (for Evaluator)
- No #REF!, #VALUE!, #NAME?, or #DIV/0! errors anywhere
- Every numeric output cell should be driven by a formula, not a hardcoded value
- Cross-sheet references must point to cells that actually exist
- Totals and subtotals must actually sum their components
- ALL numeric cells must have an explicit number format (not "General")
- Column widths must be set explicitly (not default)
- The model must look professional when opened in Excel — formatting is not optional
