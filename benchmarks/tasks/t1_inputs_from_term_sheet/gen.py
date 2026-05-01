"""Generator for t1_inputs_from_term_sheet.

Task: given a term sheet (markdown), build an `Inputs` sheet that captures
every round parameter a Series A model will need downstream.

The stub is a blank workbook. The gold is the canonical Inputs sheet
produced from the term sheet values.

Grading is per-cell numeric against gold + label existence + formula-driven
for derived fields (e.g. share price).
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- parameters extracted from term_sheet.md --------------------------------
ROUND_SIZE = 25_000_000
PRE_MONEY = 75_000_000
POST_MONEY = 100_000_000  # implied
LEAD_INVESTOR = ("Nimbus Capital", 15_000_000)
FOLLOW_INVESTORS = [
    ("Arc Ventures", 6_000_000),
    ("Orbit Partners", 4_000_000),
]
OPTION_POOL_PCT = 0.12
PRE_FD = 8_000_000
PRE_UNALLOC = 250_000
LIQ_PREF = 1.0
DIVIDEND_PCT = 0.08


def _hdr(cell) -> None:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", fgColor="DDEEFF")


def build_inputs_sheet(ws, *, fill_formulas: bool) -> None:
    ws.title = "Inputs"
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 18

    title = ws.cell(row=1, column=1, value="NorthStar Robotics — Series A Inputs")
    title.font = Font(bold=True, size=13)

    # Section 1: Round sizing
    r = 3
    s1 = ws.cell(row=r, column=1, value="Round sizing")
    _hdr(s1)
    r += 1
    ws.cell(row=r, column=1, value="Pre-money valuation")
    ws.cell(row=r, column=2, value=PRE_MONEY).number_format = "$#,##0"
    r += 1
    ws.cell(row=r, column=1, value="Round size (new money)")
    ws.cell(row=r, column=2, value=ROUND_SIZE).number_format = "$#,##0"
    r += 1
    ws.cell(row=r, column=1, value="Post-money valuation")
    if fill_formulas:
        ws.cell(row=r, column=2, value="=B4+B5").number_format = "$#,##0"
    else:
        ws.cell(row=r, column=2).number_format = "$#,##0"

    # Section 2: Investor allocations
    r += 2
    s2 = ws.cell(row=r, column=1, value="Investor allocations")
    _hdr(s2)
    r += 1
    ws.cell(row=r, column=1, value="Investor").font = Font(bold=True)
    ws.cell(row=r, column=2, value="Amount ($)").font = Font(bold=True)
    ws.cell(row=r, column=3, value="% of round").font = Font(bold=True)
    r += 1
    inv_start = r
    for name, amount in [LEAD_INVESTOR] + FOLLOW_INVESTORS:
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=amount).number_format = "$#,##0"
        if fill_formulas:
            ws.cell(row=r, column=3, value=f"=B{r}/$B$5").number_format = "0.00%"
        else:
            ws.cell(row=r, column=3).number_format = "0.00%"
        r += 1
    # Total row
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    if fill_formulas:
        ws.cell(row=r, column=2, value=f"=SUM(B{inv_start}:B{r-1})").number_format = "$#,##0"
        ws.cell(row=r, column=3, value=f"=SUM(C{inv_start}:C{r-1})").number_format = "0.00%"
    else:
        ws.cell(row=r, column=2).number_format = "$#,##0"
        ws.cell(row=r, column=3).number_format = "0.00%"

    # Section 3: Cap table mechanics
    r += 2
    s3 = ws.cell(row=r, column=1, value="Cap table mechanics")
    _hdr(s3)
    r += 1
    ws.cell(row=r, column=1, value="Pre-round fully diluted shares")
    ws.cell(row=r, column=2, value=PRE_FD).number_format = "#,##0"
    r += 1
    ws.cell(row=r, column=1, value="Pre-round unallocated option pool")
    ws.cell(row=r, column=2, value=PRE_UNALLOC).number_format = "#,##0"
    r += 1
    ws.cell(row=r, column=1, value="Target option pool % (post-money)")
    ws.cell(row=r, column=2, value=OPTION_POOL_PCT).number_format = "0.00%"

    # Section 4: Preference terms
    r += 2
    s4 = ws.cell(row=r, column=1, value="Preference terms")
    _hdr(s4)
    r += 1
    ws.cell(row=r, column=1, value="Liquidation preference (x)")
    ws.cell(row=r, column=2, value=LIQ_PREF).number_format = "0.0\"x\""
    r += 1
    ws.cell(row=r, column=1, value="Dividend rate")
    ws.cell(row=r, column=2, value=DIVIDEND_PCT).number_format = "0.0%"


def build_stub(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Inputs"
    # Blank — Builder must create the structure
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_gold(path: Path) -> None:
    wb = openpyxl.Workbook()
    build_inputs_sheet(wb.active, fill_formulas=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    recalc_xlsx(path)


def build_results_json(path: Path) -> None:
    """Gold results.json: each semantic key mapped to its Sheet!Cell address.

    Matches the layout built by build_inputs_sheet with fill_formulas=True:
      B4  Pre-money valuation
      B5  Round size
      B6  Post-money valuation (formula)
      B10 Nimbus Capital amount
      B11 Arc Ventures amount
      B12 Orbit Partners amount
      B13 Investor total (formula)
      B16 Pre-round fully diluted
      B17 Pre-round unallocated pool
      B18 Option pool target %
      B21 Liquidation preference
      B22 Dividend rate
    """
    import json
    data = {
        "pre_money_valuation": "Inputs!B4",
        "round_size": "Inputs!B5",
        "post_money_valuation": "Inputs!B6",
        "nimbus_capital_amount": "Inputs!B10",
        "arc_ventures_amount": "Inputs!B11",
        "orbit_partners_amount": "Inputs!B12",
        "investor_total": "Inputs!B13",
        "pre_round_fd": "Inputs!B16",
        "pre_round_unallocated": "Inputs!B17",
        "option_pool_target_pct": "Inputs!B18",
        "liquidation_preference": "Inputs!B21",
        "dividend_rate": "Inputs!B22",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2))


def main() -> None:
    build_stub(TASK_DIR / "stubs" / "starter.xlsx")
    build_gold(TASK_DIR / "gold" / "model.xlsx")
    build_results_json(TASK_DIR / "gold" / "results.json")
    print("t1 artifacts written.")


if __name__ == "__main__":
    main()
