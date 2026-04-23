# Revenue build — assumptions

## Annual totals
- Total annual units: 120,000
- Base unit price: $50

## Seasonality (fraction of annual units per month; must sum to 1.0)
| Month | Fraction |
|-------|----------|
| Jan   | 0.05 |
| Feb   | 0.05 |
| Mar   | 0.07 |
| Apr   | 0.08 |
| May   | 0.08 |
| Jun   | 0.09 |
| Jul   | 0.10 |
| Aug   | 0.10 |
| Sep   | 0.09 |
| Oct   | 0.10 |
| Nov   | 0.11 |
| Dec   | 0.08 |

## Price escalator
- Price is flat at $50 all 12 months (no escalation).

## Layout (on sheet "Revenue Build")
- A1: "Revenue Build" (bold, size 13)
- Row 3: headers: Month | Seasonality % | Units | Price | Revenue
- Rows 4-15: Jan through Dec
- Row 16: Totals — SUM of units, weighted-avg price (= total rev / total units), and total revenue

## Formulas
- Units_month = annual_units × seasonality
- Revenue_month = units_month × price
- Total revenue should equal annual_units × price = 120,000 × 50 = $6,000,000
