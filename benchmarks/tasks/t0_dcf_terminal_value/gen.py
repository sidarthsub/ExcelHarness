"""Generator for t0_dcf_terminal_value.

Deterministically produces:
  stubs/starter.xlsx  - workbook with Calc sheet inputs filled, answers blank
  gold/model.xlsx     - same as stub but with the correct formulas
  gold/results.json   - maps semantic keys to Calc!Bxx addresses
  gold/grading.yaml   - rubric (hardcoded, with computed expected values)

Math: Gordon-growth terminal value + PV pull-back.
  TV       = FCF_Y5 * (1 + g) / (WACC - g)
  factor   = 1 / (1 + WACC)^horizon
  PV of TV = TV * factor

Parameters:
  Y5 FCF:    10,000,000
  growth g:  0.025
  WACC:      0.09
  horizon:   5

Expected:
  TV     ~= 157,692,307.69
  factor ~= 0.649931
  PV     ~= 102,486,...
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

from benchmarks.recalc import recalc_xlsx

TASK_DIR = Path(__file__).resolve().parent

# --- parameters -------------------------------------------------------------
Y5_FCF = 10_000_000
GROWTH_G = 0.025
WACC = 0.09
HORIZON = 5

# Expected values, computed at gen time so grading.yaml stays in sync.
GOLD_TV = Y5_FCF * (1 + GROWTH_G) / (WACC - GROWTH_G)
GOLD_FACTOR = 1 / (1 + WACC) ** HORIZON
GOLD_PV = GOLD_TV * GOLD_FACTOR


# --- builders ---------------------------------------------------------------


def build_starter(path: Path, *, fill_answer: bool) -> None:
    """Build a workbook with a single Calc sheet.

    Layout:
      A1  "Terminal value (Gordon growth)"
      A3  "Year 5 FCF"              B3  10,000,000
      A4  "Terminal growth"         B4  0.025
      A5  "WACC"                    B5  0.09
      A6  "Horizon"                 B6  5
      A9  "Terminal value (Y5)"     B9  <formula if fill_answer>
      A10 "Discount factor Y5"      B10 <formula if fill_answer>
      A11 "PV of terminal value"    B11 <formula if fill_answer>
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Calc"

    hdr = Font(bold=True)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 20

    ws.cell(row=1, column=1, value="Terminal value (Gordon growth)").font = hdr

    ws.cell(row=3, column=1, value="Year 5 FCF")
    ws.cell(row=3, column=2, value=Y5_FCF).number_format = "$#,##0"

    ws.cell(row=4, column=1, value="Terminal growth")
    ws.cell(row=4, column=2, value=GROWTH_G).number_format = "0.00%"

    ws.cell(row=5, column=1, value="WACC")
    ws.cell(row=5, column=2, value=WACC).number_format = "0.00%"

    ws.cell(row=6, column=1, value="Horizon")
    ws.cell(row=6, column=2, value=HORIZON).number_format = "0"

    ws.cell(row=9, column=1, value="Terminal value (Y5)")
    if fill_answer:
        ws.cell(row=9, column=2, value="=B3*(1+B4)/(B5-B4)")
    ws.cell(row=9, column=2).number_format = "$#,##0.00"

    ws.cell(row=10, column=1, value="Discount factor Y5")
    if fill_answer:
        ws.cell(row=10, column=2, value="=1/(1+B5)^B6")
    ws.cell(row=10, column=2).number_format = "0.000000"

    ws.cell(row=11, column=1, value="PV of terminal value")
    if fill_answer:
        ws.cell(row=11, column=2, value="=B9*B10")
    ws.cell(row=11, column=2).number_format = "$#,##0.00"

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_results_json(path: Path) -> None:
    data = {
        "terminal_value_y5": "Calc!B9",
        "discount_factor_y5": "Calc!B10",
        "pv_terminal_value": "Calc!B11",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2))


def build_grading_yaml(path: Path) -> None:
    """Emit grading.yaml with the computed expected values baked in."""
    content = f"""recalc: true
checks:
  - id: calc_sheet_exists
    type: sheet_exists
    sheet: Calc
    weight: 0.3

  - id: terminal_value_y5
    type: output_key
    key: terminal_value_y5
    expected: {GOLD_TV!r}
    tolerance_rel: 1.0e-4
    require_formula: true
    weight: 1.2
    description: "Gordon-growth terminal value at end of Y5."

  - id: discount_factor_y5
    type: output_key
    key: discount_factor_y5
    expected: {GOLD_FACTOR!r}
    tolerance_rel: 1.0e-4
    require_formula: true
    weight: 0.6
    description: "Year 5 discount factor = 1/(1+WACC)^horizon."

  - id: pv_terminal_value
    type: output_key
    key: pv_terminal_value
    expected: {GOLD_PV!r}
    tolerance_rel: 1.0e-4
    require_formula: true
    weight: 1.2
    description: "PV of terminal value today = TV x factor."
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content)


def main() -> None:
    build_starter(TASK_DIR / "stubs" / "starter.xlsx", fill_answer=False)

    gold = TASK_DIR / "gold" / "model.xlsx"
    build_starter(gold, fill_answer=True)
    recalc_xlsx(gold)

    build_results_json(TASK_DIR / "gold" / "results.json")
    build_grading_yaml(TASK_DIR / "gold" / "grading.yaml")

    print(f"GOLD_TV     = {GOLD_TV:,.4f}")
    print(f"GOLD_FACTOR = {GOLD_FACTOR:,.6f}")
    print(f"GOLD_PV     = {GOLD_PV:,.4f}")
    print("t0_dcf_terminal_value artifacts written.")


if __name__ == "__main__":
    main()
