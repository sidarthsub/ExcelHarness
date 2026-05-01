"""Generator for t2_cap_table_series_ab.

Three-sheet model:
  1. Cap Table   — pre-round cap table (10M FD total)
  2. Series A    — simple pricing, no pool change
  3. Series B    — pricing off post-A FD, 10% post-money pool top-up (pre-money
                   inclusion, net of existing unallocated)

Math (non-circular by construction):
  - pre_FD  = 10_000_000           (from Cap Table)
  - Series A:
      price_A = 40M / 10M = 4.00
      new_inv_A = 10M / 4 = 2_500_000
      post_A_FD = 10M + 2.5M = 12_500_000
  - Series B:
      pre_B_FD = post_A_FD = 12_500_000
      price_B = 100M / 12.5M = 8.00
      new_inv_B = 25M / 8 = 3_125_000
      pool top-up T solves: (unalloc + T) = 10% * (pre_B_FD + new_inv_B + T)
        => T = (0.10 * (pre_B_FD + new_inv_B) - unalloc) / (1 - 0.10)
        => T = (0.10 * 15_625_000 - 200_000) / 0.9
             = (1_562_500 - 200_000) / 0.9
             = 1_513_888.888...
      post_B_FD = 12.5M + 3.125M + 1_513_888.89 = 17_138_888.89
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

from recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- parameters -------------------------------------------------------------
PRE_FD = 10_000_000
PRE_UNALLOC = 200_000

PRE_MONEY_A = 40_000_000
ROUND_SIZE_A = 10_000_000

PRE_MONEY_B = 100_000_000
ROUND_SIZE_B = 25_000_000
POOL_PCT_B = 0.10

# Cap table rows: (holder, class, shares)
CAP_ROWS = [
    ("Founder A", "Common", 4_000_000),
    ("Founder B", "Common", 3_000_000),
    ("Seed 1", "Preferred Seed", 1_500_000),
    ("Seed 2", "Preferred Seed", 1_000_000),
    ("Options (granted)", "Options", 300_000),
    ("Options (unallocated)", "Options", PRE_UNALLOC),
]


# --- gold math --------------------------------------------------------------


def compute_gold() -> dict:
    price_a = PRE_MONEY_A / PRE_FD
    new_inv_a = ROUND_SIZE_A / price_a
    post_a_fd = PRE_FD + new_inv_a

    pre_b_fd = post_a_fd
    price_b = PRE_MONEY_B / pre_b_fd
    new_inv_b = ROUND_SIZE_B / price_b
    topup_b = (POOL_PCT_B * (pre_b_fd + new_inv_b) - PRE_UNALLOC) / (1 - POOL_PCT_B)
    post_b_fd = pre_b_fd + new_inv_b + topup_b
    return {
        "price_a": price_a,
        "new_inv_a": new_inv_a,
        "post_a_fd": post_a_fd,
        "pre_b_fd": pre_b_fd,
        "price_b": price_b,
        "new_inv_b": new_inv_b,
        "topup_b": topup_b,
        "post_b_fd": post_b_fd,
    }


GOLD = compute_gold()


# --- builders ---------------------------------------------------------------


def build_cap_table(ws, *, fill_formulas: bool) -> None:
    ws.title = "Cap Table"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 14
    headers = ["Holder", "Class", "Shares"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="E8E8E8")
    for r, (h, cls, sh) in enumerate(CAP_ROWS, 2):
        ws.cell(row=r, column=1, value=h)
        ws.cell(row=r, column=2, value=cls)
        ws.cell(row=r, column=3, value=sh).number_format = "#,##0"
    total_row = len(CAP_ROWS) + 2  # row 8 (6 rows + header)
    ws.cell(row=total_row, column=1, value="Total FD").font = Font(bold=True)
    if fill_formulas:
        ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{total_row-1})").number_format = "#,##0"
    else:
        ws.cell(row=total_row, column=3).number_format = "#,##0"


def build_series_a(ws, *, fill_formulas: bool) -> None:
    ws.title = "Series A"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 18

    ws.cell(row=1, column=1, value="Series A").font = Font(bold=True, size=13)

    # Row layout per style_reference.md (rows 3-8 inputs, 10-14 movements)
    ws.cell(row=3, column=1, value="Pre-FD (from Cap Table)")
    ws.cell(row=4, column=1, value="Pre-money")
    ws.cell(row=5, column=1, value="Round size")
    ws.cell(row=6, column=1, value="Price per share").font = Font(bold=True)
    ws.cell(row=7, column=1, value="New investor shares").font = Font(bold=True)
    ws.cell(row=8, column=1, value="Option pool top-up")
    ws.cell(row=10, column=1, value="Post-A FD").font = Font(bold=True)

    ws.cell(row=4, column=2, value=PRE_MONEY_A).number_format = "$#,##0"
    ws.cell(row=5, column=2, value=ROUND_SIZE_A).number_format = "$#,##0"
    ws.cell(row=8, column=2, value=0).number_format = "#,##0"

    if fill_formulas:
        ws.cell(row=3, column=2, value="='Cap Table'!C8").number_format = "#,##0"
        ws.cell(row=6, column=2, value="=B4/B3").number_format = "$0.0000"
        ws.cell(row=7, column=2, value="=B5/B6").number_format = "#,##0"
        ws.cell(row=10, column=2, value="=B3+B7+B8").number_format = "#,##0"
    else:
        ws.cell(row=3, column=2).number_format = "#,##0"
        ws.cell(row=6, column=2).number_format = "$0.0000"
        ws.cell(row=7, column=2).number_format = "#,##0"
        ws.cell(row=10, column=2).number_format = "#,##0"


def build_series_b(ws, *, fill_formulas: bool) -> None:
    ws.title = "Series B"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 18

    ws.cell(row=1, column=1, value="Series B").font = Font(bold=True, size=13)

    ws.cell(row=3, column=1, value="Pre-FD (from Series A)")
    ws.cell(row=4, column=1, value="Pre-money")
    ws.cell(row=5, column=1, value="Round size")
    ws.cell(row=6, column=1, value="Price per share").font = Font(bold=True)
    ws.cell(row=7, column=1, value="New investor shares").font = Font(bold=True)
    ws.cell(row=8, column=1, value="Target pool % post-B")
    ws.cell(row=9, column=1, value="Pre-round unallocated")
    ws.cell(row=10, column=1, value="Option pool top-up").font = Font(bold=True)
    ws.cell(row=12, column=1, value="Post-B FD").font = Font(bold=True)

    ws.cell(row=4, column=2, value=PRE_MONEY_B).number_format = "$#,##0"
    ws.cell(row=5, column=2, value=ROUND_SIZE_B).number_format = "$#,##0"
    ws.cell(row=8, column=2, value=POOL_PCT_B).number_format = "0.00%"

    if fill_formulas:
        ws.cell(row=3, column=2, value="='Series A'!B10").number_format = "#,##0"
        ws.cell(row=6, column=2, value="=B4/B3").number_format = "$0.0000"
        ws.cell(row=7, column=2, value="=B5/B6").number_format = "#,##0"
        ws.cell(row=9, column=2, value="='Cap Table'!C7").number_format = "#,##0"
        ws.cell(row=10, column=2, value="=(B8*(B3+B7)-B9)/(1-B8)").number_format = "#,##0"
        ws.cell(row=12, column=2, value="=B3+B7+B10").number_format = "#,##0"
    else:
        for rr, fmt in ((3, "#,##0"), (6, "$0.0000"), (7, "#,##0"),
                         (9, "#,##0"), (10, "#,##0"), (12, "#,##0")):
            ws.cell(row=rr, column=2).number_format = fmt


def build_gold(path: Path) -> None:
    wb = openpyxl.Workbook()
    build_cap_table(wb.active, fill_formulas=True)
    ws2 = wb.create_sheet(title="Series A")
    build_series_a(ws2, fill_formulas=True)
    ws3 = wb.create_sheet(title="Series B")
    build_series_b(ws3, fill_formulas=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    recalc_xlsx(path)


def build_input_cap_table(path: Path) -> None:
    wb = openpyxl.Workbook()
    build_cap_table(wb.active, fill_formulas=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    recalc_xlsx(path)


def build_round_terms(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "# Round terms\n"
        "\n"
        "## Series A\n"
        "- Pre-money: $40,000,000\n"
        "- Round size: $10,000,000\n"
        "- New investor: Voyager Capital — takes the full $10M as lead.\n"
        "- Option pool: no expansion in this round.\n"
        "\n"
        "## Series B\n"
        "- Pre-money: $100,000,000\n"
        "- Round size: $25,000,000\n"
        "- New investor: Pinnacle Partners — takes the full $25M as lead.\n"
        "- Option pool: top up to 10% of post-B fully diluted shares, pre-money inclusion (net of existing unallocated).\n"
        "\n"
        "## Conventions\n"
        "- Series A price per share = Series A pre-money / pre-round FD (simple, non-circular).\n"
        "- Series B price per share = Series B pre-money / post-Series-A FD (uses A's post-money FD as the Series B denominator).\n"
        "- Post-money FD at each round = prior FD + new investor shares + pool top-up (if any).\n"
    )


def build_style_reference(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "# Style conventions for Series A and Series B sheets\n"
        "\n"
        "## Layout per round sheet\n"
        '- A1: "Series A" or "Series B" (bold, size 13)\n'
        "- Rows 3-8: round inputs (pre-money, round size, pre-FD, headline price, etc.)\n"
        "- Rows 10-14: share movements (new investor shares, pool top-up if any, post-money FD)\n"
        "\n"
        "## Formatting\n"
        '- $ cells: "$#,##0"\n'
        '- Share counts: "#,##0"\n'
        '- Percentages: "0.00%"\n'
        '- Prices: "$0.0000"\n'
    )


def build_stub(path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.active.title = "ReadMe"
    wb.active["A1"] = "Build Cap Table, Series A, Series B per brief."
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_results_json(path: Path) -> None:
    """Gold results.json: semantic keys -> Sheet!Cell addresses."""
    import json
    data = {
        # Cap Table: total at row 8
        "preround_total_fd": "Cap Table!C8",
        # Series A: row 6 price, row 7 new inv, row 10 post-A FD
        "series_a_price": "Series A!B6",
        "series_a_new_investor_shares": "Series A!B7",
        "post_a_fd": "Series A!B10",
        # Series B: row 3 pre-B FD (cross-sheet), 6 price, 7 new inv, 10 topup, 12 post-B FD
        "series_b_pre_fd": "Series B!B3",
        "series_b_price": "Series B!B6",
        "series_b_new_investor_shares": "Series B!B7",
        "series_b_pool_topup": "Series B!B10",
        "post_b_fd": "Series B!B12",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2))


def main() -> None:
    build_input_cap_table(TASK_DIR / "inputs" / "pre_cap_table.xlsx")
    build_round_terms(TASK_DIR / "inputs" / "round_terms.md")
    build_style_reference(TASK_DIR / "inputs" / "style_reference.md")
    build_stub(TASK_DIR / "stubs" / "starter.xlsx")
    build_gold(TASK_DIR / "gold" / "model.xlsx")
    build_results_json(TASK_DIR / "gold" / "results.json")
    print("t2_cap_table_series_ab artifacts written.")
    print(f"  Series A price      = ${GOLD['price_a']:.4f}")
    print(f"  Series A new inv    = {GOLD['new_inv_a']:,.2f}")
    print(f"  Post-A FD           = {GOLD['post_a_fd']:,.2f}")
    print(f"  Series B price      = ${GOLD['price_b']:.4f}")
    print(f"  Series B new inv    = {GOLD['new_inv_b']:,.2f}")
    print(f"  Series B pool topup = {GOLD['topup_b']:,.4f}")
    print(f"  Post-B FD           = {GOLD['post_b_fd']:,.4f}")


if __name__ == "__main__":
    main()
