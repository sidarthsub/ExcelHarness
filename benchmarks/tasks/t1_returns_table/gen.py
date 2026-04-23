"""Generator for t1_returns_table.

Task: given a reference cap table xlsx and a style guide markdown, build a
single `Returns Analysis` sheet showing each holder's pro-rata payout at
5 exit valuations ($100M / $250M / $500M / $1B / $2B).

The stub is a minimal ReadMe-only workbook. The gold workbook includes an
imported `Cap Table` sheet so formulas on `Returns Analysis` can reference
share counts rather than hardcoding them.

Grading is per-cell numeric against gold via results.json dereference,
with require_formula on every output_key (pure pro-rata math must be
formula-driven).
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from benchmarks.recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- cap table (reference + mirrored sheet) ---------------------------------
HOLDERS = [
    ("Founder A", 4_000_000),
    ("Founder B", 3_000_000),
    ("Seed Investors", 2_000_000),
    ("Series A Investors", 1_500_000),
    ("Options", 500_000),
]
TOTAL_FD = sum(s for _, s in HOLDERS)  # 11,000,000

# --- exit valuations --------------------------------------------------------
EXITS = [
    ("$100M", 100_000_000),
    ("$250M", 250_000_000),
    ("$500M", 500_000_000),
    ("$1B", 1_000_000_000),
    ("$2B", 2_000_000_000),
]

# --- style helpers ----------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="DDEEFF")
THIN_TOP = Border(top=Side(style="thin"))


def _hdr_cell(cell) -> None:
    cell.font = Font(bold=True)
    cell.fill = HEADER_FILL


def build_cap_table_sheet(ws) -> None:
    """Build a 'Cap Table' worksheet with Holder/Shares columns."""
    ws.title = "Cap Table"
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 14

    ws.cell(row=1, column=1, value="Holder").font = Font(bold=True)
    ws.cell(row=1, column=2, value="Shares").font = Font(bold=True)
    for i, (name, shares) in enumerate(HOLDERS, start=2):
        ws.cell(row=i, column=1, value=name)
        ws.cell(row=i, column=2, value=shares).number_format = "#,##0"


def build_returns_sheet(ws, *, fill_formulas: bool) -> None:
    """Layout (1-indexed):
        Row 1: A1 = "Returns Analysis" (title)
        Row 2: blank
        Row 3: headers: A Holder | B Shares | C % FD | D $100M | E $250M
                        | F $500M | G $1B | H $2B
        Rows 4..8: one per holder
        Row 9: Totals
    """
    ws.title = "Returns Analysis"
    ws.column_dimensions["A"].width = 24
    for col in ("B", "C", "D", "E", "F", "G", "H"):
        ws.column_dimensions[col].width = 16

    # Title
    title = ws.cell(row=1, column=1, value="Returns Analysis")
    title.font = Font(bold=True, size=13)

    # Header row (row 3)
    headers = ["Holder", "Shares", "% FD"] + [label for label, _ in EXITS]
    for c, text in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c, value=text)
        _hdr_cell(cell)
        cell.alignment = Alignment(horizontal="center")

    # Holder rows (rows 4..8). Shares reference the Cap Table sheet.
    first_row = 4
    last_row = first_row + len(HOLDERS) - 1  # 8
    total_row = last_row + 1                 # 9

    for i, (name, _shares) in enumerate(HOLDERS):
        r = first_row + i
        ws.cell(row=r, column=1, value=name)
        if fill_formulas:
            # Shares come from the imported Cap Table sheet
            ws.cell(row=r, column=2, value=f"='Cap Table'!B{i + 2}").number_format = "#,##0"
            # % FD = shares / total_FD
            ws.cell(row=r, column=3, value=f"=B{r}/$B${total_row}").number_format = "0.00%"
            # Payouts: shares * V / total_FD
            for j, (_label, v) in enumerate(EXITS):
                col = 4 + j  # D, E, F, G, H
                col_letter = openpyxl.utils.get_column_letter(col)
                ws.cell(row=r, column=col,
                        value=f"=B{r}*{v}/$B${total_row}").number_format = "$#,##0"
        else:
            ws.cell(row=r, column=2).number_format = "#,##0"
            ws.cell(row=r, column=3).number_format = "0.00%"
            for j in range(len(EXITS)):
                col = 4 + j
                ws.cell(row=r, column=col).number_format = "$#,##0"

    # Totals row
    ws.cell(row=total_row, column=1, value="Totals").font = Font(bold=True)
    # Apply top border + bold to the whole totals row
    for c in range(1, 3 + len(EXITS) + 1):  # cols 1..8
        cell = ws.cell(row=total_row, column=c)
        cell.font = Font(bold=True)
        cell.border = THIN_TOP
    if fill_formulas:
        ws.cell(row=total_row, column=2,
                value=f"=SUM(B{first_row}:B{last_row})").number_format = "#,##0"
        ws.cell(row=total_row, column=3,
                value=f"=SUM(C{first_row}:C{last_row})").number_format = "0.00%"
        for j in range(len(EXITS)):
            col = 4 + j
            col_letter = openpyxl.utils.get_column_letter(col)
            ws.cell(row=total_row, column=col,
                    value=f"=SUM({col_letter}{first_row}:{col_letter}{last_row})"
                    ).number_format = "$#,##0"
    else:
        ws.cell(row=total_row, column=2).number_format = "#,##0"
        ws.cell(row=total_row, column=3).number_format = "0.00%"
        for j in range(len(EXITS)):
            col = 4 + j
            ws.cell(row=total_row, column=col).number_format = "$#,##0"


# --- artifact builders ------------------------------------------------------


def build_cap_table_reference(path: Path) -> None:
    """inputs/cap_table.xlsx — a standalone reference file."""
    wb = openpyxl.Workbook()
    build_cap_table_sheet(wb.active)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_stub(path: Path) -> None:
    """stubs/starter.xlsx — a ReadMe-only blank workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ReadMe"
    ws["A1"] = "Builder: create Returns Analysis sheet per brief and style guide."
    ws["A1"].font = Font(bold=True)
    ws.column_dimensions["A"].width = 80
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_gold(path: Path) -> None:
    """gold/model.xlsx — Cap Table + Returns Analysis sheets."""
    wb = openpyxl.Workbook()
    # Default sheet becomes Returns Analysis; add Cap Table first
    default = wb.active
    build_cap_table_sheet(default)
    returns_ws = wb.create_sheet("Returns Analysis")
    build_returns_sheet(returns_ws, fill_formulas=True)
    # Ensure Returns Analysis is visible first (optional nicety)
    wb.active = wb.index(returns_ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    recalc_xlsx(path)


def build_results_json(path: Path) -> None:
    """Gold results.json: semantic key -> Sheet!Cell address.

    Layout on the gold Returns Analysis sheet:
        Rows: 4=Founder A, 5=Founder B, 6=Seed, 7=Series A, 8=Options, 9=Totals
        Cols: B=Shares, C=%FD, D=$100M, E=$250M, F=$500M, G=$1B, H=$2B
    """
    import json
    data = {
        "total_fd":              "Returns Analysis!B9",
        "founder_a_pct":         "Returns Analysis!C4",
        "founder_a_payout_500m": "Returns Analysis!F4",
        "seed_payout_1b":        "Returns Analysis!G6",
        "series_a_payout_2b":    "Returns Analysis!H7",
        "total_payout_500m":     "Returns Analysis!F9",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2))


def main() -> None:
    build_cap_table_reference(TASK_DIR / "inputs" / "cap_table.xlsx")
    build_stub(TASK_DIR / "stubs" / "starter.xlsx")
    build_gold(TASK_DIR / "gold" / "model.xlsx")
    build_results_json(TASK_DIR / "gold" / "results.json")
    print("t1_returns_table artifacts written.")


if __name__ == "__main__":
    main()
