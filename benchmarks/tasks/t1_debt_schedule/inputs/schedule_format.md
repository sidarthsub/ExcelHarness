# Debt Amortization Schedule — Standard Format

## Columns (left to right)
1. **Month** (1 through 60)
2. **Beginning balance**
3. **Scheduled payment** (flat monthly, equal each period)
4. **Interest portion** = beginning balance × monthly rate
5. **Principal portion** = payment − interest
6. **Ending balance** = beginning − principal

## Formulas
- Monthly rate = annual rate / 12
- Scheduled payment = PMT(monthly_rate, n_periods, -principal)  (positive value)
- For row 1: beginning balance = initial principal
- For row t > 1: beginning balance = ending balance of t-1

## Formatting
- Column widths: Month col narrow (~8); all currency columns ~14
- Currency columns: "$#,##0"
- Month column: centered, integer format
- Header row: bold, fill #EEEEEE
- Title: "Amortization Schedule" (row 1, bold, size 13)
- Summary section below the table:
  - "Total interest paid": sum of interest column, labeled
  - "Total principal paid": sum of principal column (= initial principal)
  - "Final balance": ending balance at month 60 (≈ 0)
