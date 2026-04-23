# Task: 3-statement model for Bluebird SaaS

Build a 3-year forecast for Bluebird SaaS (FY2027–2029) with FY2026 as the
historical period. Save to `output/model.xlsx`.

## Sheets (four, named exactly)

1. **Inputs** — every driver from `inputs/drivers.md`: revenue, growth,
   COGS%, OpEx%, D&A, CapEx, interest rate, tax rate, DSO, DPO, opening
   cash, opening PP&E, debt, common stock.
2. **Income Statement** — Revenue → COGS → Gross → OpEx → EBITDA → D&A →
   EBIT → Interest → Pre-tax → Tax → Net income. Four columns: 2026, 2027,
   2028, 2029.
3. **Balance Sheet** — Assets (Cash, A/R, PP&E) and L+E (A/P, Debt, Common,
   Retained earnings) with totals. Four columns, same years.
4. **Cash Flow** — Indirect method. Net income + D&A − ΔA/R + ΔA/P =
   CF ops; − CapEx = CF investing; 0 = CF financing. Net change → Beginning
   cash → Ending cash.

## Rules

- Every derived figure must be a **formula** that traces back to `Inputs`.
  No hardcoded revenue, tax, or balance figures in the statements.
- Balance sheet must balance every period: Total assets = Total L+E.
- Ending cash on Cash Flow must tie to Cash on the Balance Sheet.
- Use the historical actuals in `inputs/historical.md` as a consistency
  check — the 2026 column should match those numbers.
- Tax = 25% × MAX(0, pre-tax income) — no tax benefit on losses.
