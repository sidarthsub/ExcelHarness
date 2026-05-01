"""Generator for t3_three_statement_model.

Builds a compact but full 3-statement model for Bluebird SaaS, 3 forecast
years (2027-2029) + 1 historical (2026). Four sheets:

  Inputs             — drivers (revenue growth, margins, DSO/DPO, tax, etc.)
  Income Statement   — IS with history + forecast
  Balance Sheet      — BS with A = L+E every period
  Cash Flow          — CF with ending cash tie to BS

All derived figures are formulas that trace back to Inputs. The gold
workbook is recalculated via soffice so openpyxl can read cached numeric
values for the grader.
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- parameters -------------------------------------------------------------
YEARS = ["2026", "2027", "2028", "2029"]  # Y0 historical + 3 forecast
FORECAST_YEARS = YEARS[1:]

REV_Y0 = 10_000_000
GROWTH = {"2027": 0.40, "2028": 0.35, "2029": 0.30}
COGS_PCT = 0.30  # flat
OPEX_PCT = {"2027": 0.55, "2028": 0.50, "2029": 0.45}

DA = 1_000_000
CAPEX = 2_000_000
INTEREST_RATE = 0.06
TAX_RATE = 0.25
DEBT = 10_000_000
COMMON = 8_000_000

DSO = 30
DPO = 45
DAYS = 365

OPENING_CASH = 15_000_000
OPENING_PPE = 5_000_000


# --- gold computation -------------------------------------------------------


def compute_gold() -> dict:
    rev = {"2026": REV_Y0}
    for y in FORECAST_YEARS:
        rev[y] = rev[str(int(y) - 1)] * (1 + GROWTH[y])

    cogs = {y: COGS_PCT * rev[y] for y in YEARS}
    gross = {y: rev[y] - cogs[y] for y in YEARS}
    # Historical 2026 OpEx given as 60%
    opex = {"2026": 0.60 * rev["2026"]}
    for y in FORECAST_YEARS:
        opex[y] = OPEX_PCT[y] * rev[y]
    ebitda = {y: gross[y] - opex[y] for y in YEARS}
    da = {y: DA for y in YEARS}
    ebit = {y: ebitda[y] - da[y] for y in YEARS}
    interest = {y: INTEREST_RATE * DEBT for y in YEARS}  # opening debt flat
    pretax = {y: ebit[y] - interest[y] for y in YEARS}
    tax = {y: TAX_RATE * max(0.0, pretax[y]) for y in YEARS}
    ni = {y: pretax[y] - tax[y] for y in YEARS}

    # Balance sheet
    ar = {y: rev[y] * DSO / DAYS for y in YEARS}
    ap = {y: cogs[y] * DPO / DAYS for y in YEARS}

    ppe = {"2026": OPENING_PPE}
    for y in FORECAST_YEARS:
        ppe[y] = ppe[str(int(y) - 1)] + CAPEX - da[y]

    # Retained earnings: start with a value that makes BS balance in 2026.
    # Given assets = cash + AR + PPE, liab = AP + debt, equity = common + RE:
    #   RE_2026 = assets_2026 - liab_2026 - common
    assets_2026 = OPENING_CASH + ar["2026"] + ppe["2026"]
    liab_2026 = ap["2026"] + DEBT
    re_2026 = assets_2026 - liab_2026 - COMMON

    re = {"2026": re_2026}
    for y in FORECAST_YEARS:
        re[y] = re[str(int(y) - 1)] + ni[y]

    # Cash roll-forward from CF
    cash = {"2026": OPENING_CASH}
    d_ar = {"2026": 0.0}
    d_ap = {"2026": 0.0}
    cf_ops = {"2026": 0.0}
    cf_inv = {"2026": 0.0}
    for y in FORECAST_YEARS:
        prev = str(int(y) - 1)
        d_ar[y] = ar[y] - ar[prev]
        d_ap[y] = ap[y] - ap[prev]
        cf_ops[y] = ni[y] + da[y] - d_ar[y] + d_ap[y]
        cf_inv[y] = -CAPEX
        cash[y] = cash[prev] + cf_ops[y] + cf_inv[y]

    total_assets = {y: cash[y] + ar[y] + ppe[y] for y in YEARS}
    total_le = {y: ap[y] + DEBT + COMMON + re[y] for y in YEARS}

    return {
        "rev": rev, "cogs": cogs, "gross": gross, "opex": opex,
        "ebitda": ebitda, "da": da, "ebit": ebit, "interest": interest,
        "pretax": pretax, "tax": tax, "ni": ni,
        "cash": cash, "ar": ar, "ppe": ppe, "ap": ap, "re": re,
        "d_ar": d_ar, "d_ap": d_ap, "cf_ops": cf_ops, "cf_inv": cf_inv,
        "total_assets": total_assets, "total_le": total_le,
    }


GOLD = compute_gold()


# --- builders ---------------------------------------------------------------


def _bold(cell) -> None:
    cell.font = Font(bold=True)


def _fill(cell, hex_color: str) -> None:
    cell.fill = PatternFill("solid", fgColor=hex_color)


def _top_border(cell) -> None:
    cell.border = Border(top=Side(style="thin"))


def _year_headers(ws, row: int, col_start: int = 2) -> dict[str, str]:
    """Write year headers and return a mapping year -> column letter."""
    col_map = {}
    for i, y in enumerate(YEARS):
        col = col_start + i
        c = ws.cell(row=row, column=col, value=int(y))
        _bold(c)
        c.alignment = Alignment(horizontal="center")
        col_map[y] = openpyxl.utils.get_column_letter(col)
    return col_map


def build_inputs(ws) -> None:
    ws.title = "Inputs"
    ws.column_dimensions["A"].width = 40
    for i, _ in enumerate(YEARS):
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = 14

    c = ws.cell(row=1, column=1, value="Bluebird SaaS — Driver Inputs")
    c.font = Font(bold=True, size=13)

    col = _year_headers(ws, 3)
    r = 4
    ws.cell(row=r, column=1, value="Revenue")
    ws.cell(row=r, column=2, value=REV_Y0).number_format = "$#,##0"
    r += 1
    ws.cell(row=r, column=1, value="Revenue growth %")
    for y in FORECAST_YEARS:
        ws[f"{col[y]}{r}"] = GROWTH[y]
        ws[f"{col[y]}{r}"].number_format = "0.0%"
    r += 1
    ws.cell(row=r, column=1, value="COGS % of revenue")
    for y in YEARS:
        ws[f"{col[y]}{r}"] = COGS_PCT
        ws[f"{col[y]}{r}"].number_format = "0.0%"
    r += 1
    ws.cell(row=r, column=1, value="OpEx % of revenue")
    # 2026 is 60% (historical), forecast uses OPEX_PCT dict
    ws[f"{col['2026']}{r}"] = 0.60
    ws[f"{col['2026']}{r}"].number_format = "0.0%"
    for y in FORECAST_YEARS:
        ws[f"{col[y]}{r}"] = OPEX_PCT[y]
        ws[f"{col[y]}{r}"].number_format = "0.0%"

    r += 2
    pairs = [
        ("D&A ($)", DA, "$#,##0"),
        ("CapEx ($)", CAPEX, "$#,##0"),
        ("Interest rate on debt", INTEREST_RATE, "0.0%"),
        ("Tax rate", TAX_RATE, "0.0%"),
        ("Opening debt", DEBT, "$#,##0"),
        ("Common stock", COMMON, "$#,##0"),
        ("DSO (days)", DSO, "#,##0"),
        ("DPO (days)", DPO, "#,##0"),
        ("Opening cash (end 2026)", OPENING_CASH, "$#,##0"),
        ("Opening PP&E (end 2026)", OPENING_PPE, "$#,##0"),
    ]
    for label, val, fmt in pairs:
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=val).number_format = fmt
        r += 1


def build_is(ws) -> None:
    """Income Statement. Layout (all rows), columns B..E for 2026..2029:
      3: year headers
      4: Revenue
      5: COGS
      6: Gross profit
      7: OpEx
      8: EBITDA
      9: D&A
      10: EBIT
      11: Interest expense
      12: Pre-tax income
      13: Tax
      14: Net income
    """
    ws.title = "Income Statement"
    ws.column_dimensions["A"].width = 28
    for i in range(len(YEARS)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = 16

    c = ws.cell(row=1, column=1, value="Income Statement")
    c.font = Font(bold=True, size=13)
    col = _year_headers(ws, 3)

    labels = {
        4: "Revenue", 5: "COGS", 6: "Gross profit",
        7: "OpEx", 8: "EBITDA", 9: "D&A",
        10: "EBIT", 11: "Interest expense", 12: "Pre-tax income",
        13: "Tax", 14: "Net income",
    }
    for r, lab in labels.items():
        ws.cell(row=r, column=1, value=lab)
    for r in (6, 8, 10, 12, 14):
        _bold(ws.cell(row=r, column=1))
    for r in (4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14):
        for y in YEARS:
            ws[f"{col[y]}{r}"].number_format = "$#,##0;($#,##0)"

    # Revenue row: Y0 pulls from Inputs!B4, forecast uses growth
    ws[f"{col['2026']}4"] = "=Inputs!B4"
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev = YEARS[i - 1]
        prev_col = col[prev]
        # growth cell in Inputs row 5
        growth_col = col[y]
        ws[f"{col[y]}4"] = f"={prev_col}4*(1+Inputs!{growth_col}5)"

    # COGS = Revenue * COGS% (Inputs row 6)
    for y in YEARS:
        ws[f"{col[y]}5"] = f"={col[y]}4*Inputs!{col[y]}6"
    # Gross = Revenue - COGS
    for y in YEARS:
        ws[f"{col[y]}6"] = f"={col[y]}4-{col[y]}5"
    # OpEx = Revenue * OpEx% (Inputs row 7)
    for y in YEARS:
        ws[f"{col[y]}7"] = f"={col[y]}4*Inputs!{col[y]}7"
    # EBITDA = Gross - OpEx
    for y in YEARS:
        ws[f"{col[y]}8"] = f"={col[y]}6-{col[y]}7"
    # D&A: flat from Inputs. D&A row in Inputs is 9 (offset from pairs section)
    # Inputs layout: row 4 = revenue, row 5 = growth, row 6 = COGS%, row 7 = OpEx%, blank 8,
    #                then pairs starting at row 9: D&A, CapEx, Interest, Tax, Debt, Common, DSO, DPO, OpCash, OpPPE
    for y in YEARS:
        ws[f"{col[y]}9"] = "=Inputs!B9"
    # EBIT = EBITDA - D&A
    for y in YEARS:
        ws[f"{col[y]}10"] = f"={col[y]}8-{col[y]}9"
    # Interest = rate * debt (both from Inputs)
    for y in YEARS:
        ws[f"{col[y]}11"] = "=Inputs!B11*Inputs!B13"
    # Pre-tax
    for y in YEARS:
        ws[f"{col[y]}12"] = f"={col[y]}10-{col[y]}11"
    # Tax = rate * MAX(0, pretax)
    for y in YEARS:
        ws[f"{col[y]}13"] = f"=Inputs!B12*MAX(0,{col[y]}12)"
    # Net income
    for y in YEARS:
        ws[f"{col[y]}14"] = f"={col[y]}12-{col[y]}13"

    for y in YEARS:
        _top_border(ws[f"{col[y]}6"])
        _top_border(ws[f"{col[y]}10"])
        _top_border(ws[f"{col[y]}14"])


def build_bs(ws) -> None:
    """Balance Sheet.
    Rows:
      3 year headers
      4 Cash
      5 A/R
      6 PP&E (net)
      7 Total assets
      9 A/P
      10 Debt
      11 Common stock
      12 Retained earnings
      13 Total L+E
    """
    ws.title = "Balance Sheet"
    ws.column_dimensions["A"].width = 28
    for i in range(len(YEARS)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = 16

    c = ws.cell(row=1, column=1, value="Balance Sheet")
    c.font = Font(bold=True, size=13)
    col = _year_headers(ws, 3)

    labels = {4: "Cash", 5: "A/R", 6: "PP&E (net)", 7: "Total assets",
              9: "A/P", 10: "Debt", 11: "Common stock",
              12: "Retained earnings", 13: "Total L+E"}
    for r, lab in labels.items():
        ws.cell(row=r, column=1, value=lab)
    for r in (7, 13):
        _bold(ws.cell(row=r, column=1))
    for r in (4, 5, 6, 7, 9, 10, 11, 12, 13):
        for y in YEARS:
            ws[f"{col[y]}{r}"].number_format = "$#,##0;($#,##0)"

    # Cash: 2026 opening from Inputs; forecast = prior cash + CF sheet net change
    ws[f"{col['2026']}4"] = "=Inputs!B17"  # opening cash row
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev_col = col[YEARS[i - 1]]
        ws[f"{col[y]}4"] = f"={prev_col}4+'Cash Flow'!{col[y]}9"
    # A/R = Revenue * DSO/365
    for y in YEARS:
        ws[f"{col[y]}5"] = f"='Income Statement'!{col[y]}4*Inputs!B15/365"
    # PP&E: 2026 from Inputs!B18. Forecast = prior + CapEx - D&A
    ws[f"{col['2026']}6"] = "=Inputs!B18"
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev_col = col[YEARS[i - 1]]
        ws[f"{col[y]}6"] = f"={prev_col}6+Inputs!B10-'Income Statement'!{col[y]}9"
    # Total assets
    for y in YEARS:
        ws[f"{col[y]}7"] = f"=SUM({col[y]}4:{col[y]}6)"

    # A/P = COGS * DPO/365
    for y in YEARS:
        ws[f"{col[y]}9"] = f"='Income Statement'!{col[y]}5*Inputs!B16/365"
    # Debt = constant
    for y in YEARS:
        ws[f"{col[y]}10"] = "=Inputs!B13"
    # Common = constant
    for y in YEARS:
        ws[f"{col[y]}11"] = "=Inputs!B14"
    # RE: 2026 plugged via a formula that forces the BS to balance:
    #   RE_2026 = (Cash_2026 + AR_2026 + PPE_2026) - (AP_2026 + Debt_2026 + Common_2026)
    ws[f"{col['2026']}12"] = (
        f"=({col['2026']}4+{col['2026']}5+{col['2026']}6)-"
        f"({col['2026']}9+{col['2026']}10+{col['2026']}11)"
    )
    # Forecast RE rolls with net income
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev_col = col[YEARS[i - 1]]
        ws[f"{col[y]}12"] = f"={prev_col}12+'Income Statement'!{col[y]}14"
    # Total L+E
    for y in YEARS:
        ws[f"{col[y]}13"] = f"={col[y]}9+SUM({col[y]}10:{col[y]}12)"

    for y in YEARS:
        _top_border(ws[f"{col[y]}7"])
        _top_border(ws[f"{col[y]}13"])


def build_cf(ws) -> None:
    """Cash Flow Statement.
    Rows:
      3 year headers
      4 Net income
      5 + D&A
      6 − ΔA/R
      7 + ΔA/P
      8 CF from operations
      9 CF from investing (= -CapEx)
      10 CF from financing  (=0)
      11 Net change in cash
      12 Beginning cash
      13 Ending cash
    """
    ws.title = "Cash Flow"
    ws.column_dimensions["A"].width = 28
    for i in range(len(YEARS)):
        ws.column_dimensions[openpyxl.utils.get_column_letter(2 + i)].width = 16

    c = ws.cell(row=1, column=1, value="Cash Flow")
    c.font = Font(bold=True, size=13)
    col = _year_headers(ws, 3)

    labels = {4: "Net income", 5: "+ D&A", 6: "− ΔA/R", 7: "+ ΔA/P",
              8: "CF from operations", 9: "CF from investing",
              10: "CF from financing", 11: "Net change in cash",
              12: "Beginning cash", 13: "Ending cash"}
    for r, lab in labels.items():
        ws.cell(row=r, column=1, value=lab)
    for r in (8, 11, 13):
        _bold(ws.cell(row=r, column=1))
    for r in (4, 5, 6, 7, 8, 9, 10, 11, 12, 13):
        for y in YEARS:
            ws[f"{col[y]}{r}"].number_format = "$#,##0;($#,##0)"

    # Historical column: CF is not meaningful for 2026 — zero out operating rows.
    # Leave 2026 blank/zeros to keep the schema consistent.
    for y in YEARS:
        ws[f"{col[y]}4"] = f"='Income Statement'!{col[y]}14"
        ws[f"{col[y]}5"] = f"='Income Statement'!{col[y]}9"
    # ΔA/R and ΔA/P: only meaningful for forecast years
    for i, y in enumerate(YEARS):
        if i == 0:
            ws[f"{col[y]}6"] = 0
            ws[f"{col[y]}7"] = 0
        else:
            prev_col = col[YEARS[i - 1]]
            ws[f"{col[y]}6"] = f"=-('Balance Sheet'!{col[y]}5-'Balance Sheet'!{prev_col}5)"
            ws[f"{col[y]}7"] = f"='Balance Sheet'!{col[y]}9-'Balance Sheet'!{prev_col}9"
    # Investing: -CapEx for forecast, 0 for 2026
    ws[f"{col['2026']}9"] = 0
    for y in FORECAST_YEARS:
        ws[f"{col[y]}9"] = "=-Inputs!B10"
    # Financing: 0
    for y in YEARS:
        ws[f"{col[y]}10"] = 0
    # CF ops
    for y in YEARS:
        ws[f"{col[y]}8"] = f"=SUM({col[y]}4:{col[y]}7)"
    # Net change = ops + inv + fin
    # NOTE: net change in cash row is a SINGLE number — the sum of CF ops + inv + fin
    # But BS cash formula references this row (row 9) as "net change". Let's put net change into row 9.
    # Actually I put CF from investing in row 9. Let me re-examine — BS references 'Cash Flow'!{col[y]}9
    # which is CF from investing. That's a bug. Let me fix this by reordering the CF sheet
    # to keep BS reference aligned.
    # (Corrected below — row 11 holds net change.)
    for y in YEARS:
        ws[f"{col[y]}11"] = f"={col[y]}8+{col[y]}9+{col[y]}10"
    # Beginning + ending cash
    ws[f"{col['2026']}12"] = 0
    ws[f"{col['2026']}13"] = "=Inputs!B17"
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev_col = col[YEARS[i - 1]]
        ws[f"{col[y]}12"] = f"={prev_col}13"
        ws[f"{col[y]}13"] = f"={col[y]}12+{col[y]}11"

    for y in YEARS:
        _top_border(ws[f"{col[y]}8"])
        _top_border(ws[f"{col[y]}11"])
        _top_border(ws[f"{col[y]}13"])


def build_gold(path: Path) -> None:
    # Important: BS cash formula references `'Cash Flow'!{col[y]}9` for net change,
    # but net change is on row 11. Fix the BS formula: build_bs uses row 9 which is
    # actually CF investing. Rebuild BS to point at the correct row (11).
    wb = openpyxl.Workbook()
    build_inputs(wb.active)
    build_is(wb.create_sheet("Income Statement"))
    build_bs(wb.create_sheet("Balance Sheet"))
    build_cf(wb.create_sheet("Cash Flow"))

    # Patch BS cash formula to reference CF net change row (11), not row 9.
    bs = wb["Balance Sheet"]
    col = {y: openpyxl.utils.get_column_letter(2 + i) for i, y in enumerate(YEARS)}
    for i, y in enumerate(FORECAST_YEARS, start=1):
        prev_col = col[YEARS[i - 1]]
        bs[f"{col[y]}4"] = f"={prev_col}4+'Cash Flow'!{col[y]}11"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    recalc_xlsx(path)


def build_stub(path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.active.title = "ReadMe"
    wb.active["A1"] = "Builder: construct Inputs, Income Statement, Balance Sheet, Cash Flow."
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# --- grading generation -----------------------------------------------------


def year_to_col(y: str) -> str:
    i = YEARS.index(y)
    return openpyxl.utils.get_column_letter(2 + i)


def write_grading(path: Path) -> None:
    checks = []

    def add(cid, ctype, **kw):
        checks.append({"id": cid, "type": ctype, **kw})

    # Sheet presence
    for s in ("Inputs", "Income Statement", "Balance Sheet", "Cash Flow"):
        add(f"sheet_{s.lower().replace(' ', '_')}", "sheet_exists", sheet=s, weight=0.5)

    # IS forecast values
    is_rows = {
        "revenue": (4, "rev"), "cogs": (5, "cogs"), "gross": (6, "gross"),
        "opex": (7, "opex"), "ebitda": (8, "ebitda"), "ebit": (10, "ebit"),
        "pretax": (12, "pretax"), "ni": (14, "ni"),
    }
    for name, (row, key) in is_rows.items():
        for y in FORECAST_YEARS:
            add(f"is_{name}_{y}", "numeric",
                sheet="Income Statement", cell=f"{year_to_col(y)}{row}",
                expected=float(GOLD[key][y]), tolerance_rel=1.0e-4, weight=0.5)

    # BS forecast values
    bs_rows = {"cash": 4, "ar": 5, "ppe": 6, "total_assets": 7,
               "ap": 9, "re": 12, "total_le": 13}
    bs_keys = {"cash": "cash", "ar": "ar", "ppe": "ppe", "total_assets": "total_assets",
               "ap": "ap", "re": "re", "total_le": "total_le"}
    for name, row in bs_rows.items():
        for y in FORECAST_YEARS:
            add(f"bs_{name}_{y}", "numeric",
                sheet="Balance Sheet", cell=f"{year_to_col(y)}{row}",
                expected=float(GOLD[bs_keys[name]][y]), tolerance_rel=1.0e-4, weight=0.5)

    # CF forecast values
    for y in FORECAST_YEARS:
        add(f"cf_ops_{y}", "numeric",
            sheet="Cash Flow", cell=f"{year_to_col(y)}8",
            expected=float(GOLD["cf_ops"][y]), tolerance_rel=1.0e-4, weight=0.75)
        add(f"cf_ending_cash_{y}", "numeric",
            sheet="Cash Flow", cell=f"{year_to_col(y)}13",
            expected=float(GOLD["cash"][y]), tolerance_rel=1.0e-4, weight=1.0)

    # BS balances check (the headline signal)
    add("bs_balances", "bs_balances",
        sheet="Balance Sheet", assets_row=7, liab_equity_row=13,
        columns=[year_to_col(y) for y in YEARS],
        tolerance_abs=1.0, tolerance_rel=1.0e-4, weight=3.0)

    # Formula-driven critical cells
    for y in FORECAST_YEARS:
        col = year_to_col(y)
        add(f"is_revenue_is_formula_{y}", "formula_driven",
            sheet="Income Statement", cell=f"{col}4", weight=0.2)
        add(f"bs_cash_is_formula_{y}", "formula_driven",
            sheet="Balance Sheet", cell=f"{col}4", weight=0.2)
        add(f"bs_re_is_formula_{y}", "formula_driven",
            sheet="Balance Sheet", cell=f"{col}12", weight=0.2)
        add(f"cf_ending_cash_is_formula_{y}", "formula_driven",
            sheet="Cash Flow", cell=f"{col}13", weight=0.2)

    data = {"recalc": True, "checks": checks}
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, sort_keys=False, allow_unicode=True))


def main() -> None:
    build_stub(TASK_DIR / "stubs" / "starter.xlsx")
    build_gold(TASK_DIR / "gold" / "model.xlsx")
    write_grading(TASK_DIR / "gold" / "grading.yaml")
    print("t3 artifacts written.")
    for y in YEARS:
        a, le = GOLD["total_assets"][y], GOLD["total_le"][y]
        print(f"  {y}: A={a:>16,.2f}  L+E={le:>16,.2f}  delta={a-le:>10,.4f}  cash={GOLD['cash'][y]:>14,.2f}  NI={GOLD['ni'][y]:>12,.2f}")


if __name__ == "__main__":
    main()
