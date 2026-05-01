"""Generator for t0_option_pool_issuance.

Deterministically produces:
  inputs/pre_cap_table.xlsx   — the reference cap table (pre-round)
  stubs/starter.xlsx          — workbook with Inputs sheet filled, answer cell blank
  gold/model.xlsx             — same as stub but with the correct answer

Math: option pool expansion under pre-money inclusion.
  new_pool_shares = (pool_pct_post_money * post_money_FD - existing_unallocated) / (1 - pool_pct_post_money)

  where post_money_FD = existing_FD + new_investor_shares + new_pool_shares
  (solve for new_pool_shares; under pre-money inclusion new investor shares are
  fixed by investment_$ / share_price and share_price is determined from
  pre_money / (existing_FD + new_pool_shares - existing_unallocated) ... )

To keep Tier-0 atomic, we fix investor shares so the problem reduces to:
  target pool = 10% of (existing_FD + issued_pool_new - existing_unallocated + investor_shares)

Given:
  existing_FD     = 10,000,000
  existing_unalloc = 500,000
  investor_shares = 2,500,000   (already determined; Series A closed)
  target_pool_pct = 10%

Solve:
  0.10 * (existing_FD + investor_shares + new_pool) = existing_unalloc + new_pool
  -> new_pool = (0.10 * (existing_FD + investor_shares) - existing_unalloc) / 0.90
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- parameters -------------------------------------------------------------
EXISTING_FD = 10_000_000
EXISTING_UNALLOC = 500_000
INVESTOR_SHARES = 2_500_000
TARGET_POOL_PCT = 0.10

GOLD_NEW_POOL = (TARGET_POOL_PCT * (EXISTING_FD + INVESTOR_SHARES) - EXISTING_UNALLOC) / (1 - TARGET_POOL_PCT)
# plug-in: 1_944_444.44...


# --- builders ---------------------------------------------------------------


def build_input_cap_table(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cap Table"
    hdr_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="E8E8E8")

    headers = ["Holder", "Class", "Shares", "% FD"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hdr_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center")

    rows = [
        ["Founder A", "Common", 4_000_000],
        ["Founder B", "Common", 3_000_000],
        ["Seed Investor 1", "Preferred Seed", 1_500_000],
        ["Seed Investor 2", "Preferred Seed", 1_000_000],
        ["Option Pool (allocated)", "Options (granted)", 500_000],
        ["Option Pool (unallocated)", "Options (unallocated)", EXISTING_UNALLOC],
    ]
    for r, (h, cls, sh) in enumerate(rows, 2):
        ws.cell(row=r, column=1, value=h)
        ws.cell(row=r, column=2, value=cls)
        ws.cell(row=r, column=3, value=sh).number_format = "#,##0"
        ws.cell(row=r, column=4, value=f"=C{r}/C${len(rows)+2}").number_format = "0.00%"

    total_row = len(rows) + 2
    ws.cell(row=total_row, column=1, value="Total FD").font = hdr_font
    ws.cell(row=total_row, column=3, value=f"=SUM(C2:C{total_row-1})").font = hdr_font
    ws.cell(row=total_row, column=3).number_format = "#,##0"
    ws.cell(row=total_row, column=4, value=1.0).number_format = "0.00%"

    for col, w in [("A", 28), ("B", 22), ("C", 14), ("D", 10)]:
        ws.column_dimensions[col].width = w

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_starter(path: Path, *, fill_answer: float | None = None) -> None:
    """Build a workbook with an Inputs sheet + Calc sheet.

    The Calc sheet holds:
      B2  Existing FD (hardcoded link-worthy value)
      B3  Existing unallocated
      B4  New investor shares
      B5  Target pool % post-money
      B10 Answer: new option pool shares to issue

    If fill_answer is provided, B10 is populated. Otherwise left blank.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"

    hdr = Font(bold=True)
    label_col_width = 46
    ws.column_dimensions["A"].width = label_col_width
    ws.column_dimensions["B"].width = 18

    ws.cell(row=1, column=1, value="Option pool top-up (pre-money inclusion)").font = hdr

    ws.cell(row=2, column=1, value="Existing fully diluted shares")
    ws.cell(row=2, column=2, value=EXISTING_FD).number_format = "#,##0"

    ws.cell(row=3, column=1, value="Existing unallocated options")
    ws.cell(row=3, column=2, value=EXISTING_UNALLOC).number_format = "#,##0"

    ws.cell(row=4, column=1, value="New investor shares (already determined)")
    ws.cell(row=4, column=2, value=INVESTOR_SHARES).number_format = "#,##0"

    ws.cell(row=5, column=1, value="Target option pool % (post-money)")
    ws.cell(row=5, column=2, value=TARGET_POOL_PCT).number_format = "0.00%"

    ws.cell(row=9, column=1, value="Answer:").font = hdr
    ws.cell(row=10, column=1, value="New option pool shares to issue")
    if fill_answer is not None:
        # Write as formula referencing the inputs so the grader's hardcode check passes.
        ws.cell(row=10, column=2, value="=(B5*(B2+B4)-B3)/(1-B5)")
    ws.cell(row=10, column=2).number_format = "#,##0"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main() -> None:
    inp = TASK_DIR / "inputs" / "pre_cap_table.xlsx"
    build_input_cap_table(inp)
    recalc_xlsx(inp)  # populate cached values for the SUM / percentage formulas

    build_starter(TASK_DIR / "stubs" / "starter.xlsx", fill_answer=None)
    gold = TASK_DIR / "gold" / "model.xlsx"
    build_starter(gold, fill_answer=GOLD_NEW_POOL)
    recalc_xlsx(gold)
    print(f"GOLD_NEW_POOL = {GOLD_NEW_POOL:,.4f}")


if __name__ == "__main__":
    main()
