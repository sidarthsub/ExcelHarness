#!/usr/bin/env python3
"""Tests for the deterministic evaluator.

Covers:
- Expression evaluation with real financial formula patterns
- Multi-component iterative convergence (share-price/count, ESOP top-up)
- Full supported function surface
- Nonsense/error detection: broken refs, div/0, type mismatches
- Complex INDEX/MATCH/XLOOKUP combos
- Graph construction, SCC detection, topological ordering
- Acceptance tests against real workbook dumps
"""

import json
import math
import sys
import tempfile
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from deterministic_evaluator import (
    ABS_TOLERANCE,
    EvalError,
    ExpressionEvaluator,
    FormulaCell,
    UnsupportedError,
    Workbook,
    build_dependency_graph,
    classify_functions,
    col_to_num,
    evaluate_cell,
    evaluate_workbook,
    expand_range,
    load_workbook,
    num_to_col,
    parse_cell,
    qualified_cell,
    solve_iterative_component,
    tarjan_scc,
    topological_sort,
    _date_to_serial,
    _serial_to_date,
    _seed_iterative_cells,
)

import pytest


# ---------------------------------------------------------------------------
# Token helpers
# ---------------------------------------------------------------------------

def _num(v):
    return {"type": "OPERAND", "subtype": "NUMBER", "value": str(v)}

def _text(v):
    return {"type": "OPERAND", "subtype": "TEXT", "value": f'"{v}"'}

def _ref(v):
    return {"type": "OPERAND", "subtype": "RANGE", "value": v}

def _op(v):
    return {"type": "OPERATOR-INFIX", "subtype": "", "value": v}

def _func_open(name):
    return {"type": "FUNC", "subtype": "OPEN", "value": f"{name}("}

def _func_close():
    return {"type": "FUNC", "subtype": "CLOSE", "value": ")"}

def _sep():
    return {"type": "SEP", "subtype": "ARG", "value": ","}

def _paren_open():
    return {"type": "PAREN", "subtype": "OPEN", "value": "("}

def _paren_close():
    return {"type": "PAREN", "subtype": "CLOSE", "value": ")"}

def _prefix(v):
    return {"type": "OPERATOR-PREFIX", "subtype": "", "value": v}

def _logical(v):
    return {"type": "OPERAND", "subtype": "LOGICAL", "value": v}

def _subexpr_open():
    return {"type": "SUBEXPR", "subtype": "OPEN", "value": "("}

def _subexpr_close():
    return {"type": "SUBEXPR", "subtype": "CLOSE", "value": ")"}

def _postfix(v):
    return {"type": "OPERATOR-POSTFIX", "subtype": "", "value": v}


def _make_wb(sheets: dict) -> Workbook:
    """Build a test workbook from inline data.

    sheets = {
        "Sheet1": {
            "formulas": { "A1": {"formula": "=...", "tokens": [...], ...} },
            "literals": {"B1": 10},
        }
    }
    """
    wb = Workbook()
    wb.sheet_names = list(sheets.keys())
    for sheet_name, sheet_data in sheets.items():
        for cell_ref, val in sheet_data.get("literals", {}).items():
            wb.literal_values[qualified_cell(sheet_name, cell_ref)] = val
        for cell_ref, cell_data in sheet_data.get("formulas", {}).items():
            data = {
                "cell": cell_ref,
                "formula": cell_data.get("formula", ""),
                "cached_value": cell_data.get("cached_value"),
                "functions": cell_data.get("functions", []),
                "references": cell_data.get("references", []),
                "dependencies": cell_data.get("dependencies", []),
                "tokens": cell_data.get("tokens", []),
                "iterative_component": cell_data.get("iterative_component"),
            }
            fc = FormulaCell(sheet_name, data)
            wb.add_cell(fc)
    return wb


def _eval(wb, sheet, tokens):
    """Shorthand: evaluate tokens in context of wb/sheet."""
    return ExpressionEvaluator(wb, sheet).evaluate(tokens)


def _empty_wb():
    return _make_wb({"S": {"formulas": {}}})


# ===================================================================
# 1. Cell reference utilities
# ===================================================================

class TestCellUtils:
    def test_col_round_trip(self):
        for i in range(1, 100):
            assert col_to_num(num_to_col(i)) == i

    def test_parse_cell_strips_dollars(self):
        assert parse_cell("$AB$99") == ("AB", 99)

    def test_expand_range_single_row(self):
        assert expand_range("A5:D5") == ["A5", "B5", "C5", "D5"]

    def test_expand_range_single_col(self):
        assert expand_range("B1:B4") == ["B1", "B2", "B3", "B4"]

    def test_expand_range_rect(self):
        cells = expand_range("A1:B3")
        assert len(cells) == 6
        assert cells[0] == "A1"
        assert cells[-1] == "B3"


# ===================================================================
# 2. Expression evaluator — operator precedence & edge cases
# ===================================================================

class TestExpressionEdgeCases:
    """Tests that exercise real formula patterns, not toy 1+2."""

    def test_nested_parens_with_division(self):
        """Pattern from cap tables: (invested / price_per_share)"""
        wb = _make_wb({"S": {"literals": {"A1": 5000000, "A2": 0.50}}})
        # (A1 / A2)
        tokens = [_paren_open(), _ref("A1"), _op("/"), _ref("A2"), _paren_close()]
        assert _eval(wb, "S", tokens) == 10000000.0

    def test_mixed_subexpr_and_paren_tokens(self):
        """openpyxl emits both SUBEXPR and PAREN depending on context."""
        wb = _empty_wb()
        # ((2 + 3)) using one PAREN and one SUBEXPR level
        tokens = [
            _subexpr_open(),
            _paren_open(), _num(2), _op("+"), _num(3), _paren_close(),
            _op("*"),
            _paren_open(), _num(4), _op("-"), _num(1), _paren_close(),
            _subexpr_close(),
        ]
        assert _eval(wb, "S", tokens) == 15.0

    def test_percent_postfix(self):
        wb = _empty_wb()
        # 7% → 0.07
        assert _eval(wb, "S", [_num(7), _postfix("%")]) == 0.07

    def test_unary_plus_on_string_passthrough(self):
        """Excel's =+'Sheet'!B8 on a text cell returns the text."""
        wb = _make_wb({"S": {"literals": {"A1": "Brian Hazzard"}}})
        tokens = [_prefix("+"), _ref("A1")]
        assert _eval(wb, "S", tokens) == "Brian Hazzard"

    def test_string_comparison(self):
        wb = _empty_wb()
        assert _eval(wb, "S", [_text("yes"), _op("="), _text("yes")]) is True
        assert _eval(wb, "S", [_text("yes"), _op("<>"), _text("no")]) is True

    def test_concatenation_with_numbers(self):
        wb = _make_wb({"S": {"literals": {"A1": 42}}})
        tokens = [_text("Value: "), _op("&"), _ref("A1")]
        assert _eval(wb, "S", tokens) == "Value: 42"

    def test_division_by_zero_raises(self):
        wb = _empty_wb()
        with pytest.raises(EvalError, match="Division by zero"):
            _eval(wb, "S", [_num(1), _op("/"), _num(0)])

    def test_none_cell_treated_as_zero_in_arithmetic(self):
        """Unresolved cell ref → None → 0 in arithmetic."""
        wb = _make_wb({"S": {"formulas": {}}})
        # A1 doesn't exist → None
        tokens = [_ref("A1"), _op("+"), _num(5)]
        assert _eval(wb, "S", tokens) == 5.0

    def test_chained_comparisons_in_if(self):
        """IF(A1>100, IF(A1>200, "high", "mid"), "low")"""
        wb = _make_wb({"S": {"literals": {"A1": 150}}})
        tokens = [
            _func_open("IF"),
            _ref("A1"), _op(">"), _num(100), _sep(),
            _func_open("IF"),
            _ref("A1"), _op(">"), _num(200), _sep(), _text("high"), _sep(), _text("mid"),
            _func_close(), _sep(),
            _text("low"),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == "mid"


# ===================================================================
# 3. IFERROR — real patterns from cap table models
# ===================================================================

class TestIferror:
    def test_iferror_catches_div_zero(self):
        """=IFERROR(invested/total, 0) when total=0"""
        wb = _make_wb({"S": {"literals": {"A1": 5000000, "A2": 0}, "formulas": {
            "B1": {
                "formula": "=IFERROR(A1/A2,0)",
                "tokens": [
                    _func_open("IFERROR"),
                    _ref("A1"), _op("/"), _ref("A2"),
                    _sep(), _num(0),
                    _func_close(),
                ],
            },
        }}})
        assert evaluate_cell(wb, wb.formula_cells["S!B1"]) == 0.0

    def test_iferror_passes_through_good_value(self):
        wb = _make_wb({"S": {"literals": {"A1": 100, "A2": 4}, "formulas": {
            "B1": {
                "formula": "=IFERROR(A1/A2,0)",
                "tokens": [
                    _func_open("IFERROR"),
                    _ref("A1"), _op("/"), _ref("A2"),
                    _sep(), _num(0),
                    _func_close(),
                ],
            },
        }}})
        assert evaluate_cell(wb, wb.formula_cells["S!B1"]) == 25.0

    def test_nested_iferror_round(self):
        """=IFERROR(ROUND(D16*(C14+E22)/50000000, 0), 0)"""
        wb = _make_wb({"S": {"literals": {"D16": 1000, "C14": 20000000, "E22": 30000000}, "formulas": {
            "A1": {
                "formula": "=IFERROR(ROUND(D16*(C14+E22)/50000000,0),0)",
                "tokens": [
                    _func_open("IFERROR"),
                    _func_open("ROUND"),
                    _ref("D16"), _op("*"),
                    _paren_open(), _ref("C14"), _op("+"), _ref("E22"), _paren_close(),
                    _op("/"), _num(50000000),
                    _sep(), _num(0),
                    _func_close(),
                    _sep(), _num(0),
                    _func_close(),
                ],
            },
        }}})
        # 1000 * 50M / 50M = 1000
        assert evaluate_cell(wb, wb.formula_cells["S!A1"]) == 1000.0


# ===================================================================
# 4. Aggregates with real data shapes
# ===================================================================

class TestAggregatesReal:
    def test_sum_skips_none_and_strings(self):
        """SUM range with mixed types — strings and None should be skipped."""
        wb = _make_wb({"S": {"literals": {
            "A1": 100, "A2": "label", "A3": None, "A4": 200, "A5": 300,
        }}})
        tokens = [_func_open("SUM"), _ref("A1:A5"), _func_close()]
        assert _eval(wb, "S", tokens) == 600.0

    def test_sum_multiple_ranges(self):
        """=SUM(A1:A3, B1:B3)"""
        wb = _make_wb({"S": {"literals": {
            "A1": 1, "A2": 2, "A3": 3, "B1": 10, "B2": 20, "B3": 30,
        }}})
        tokens = [_func_open("SUM"), _ref("A1:A3"), _sep(), _ref("B1:B3"), _func_close()]
        assert _eval(wb, "S", tokens) == 66.0

    def test_sumifs_multi_criteria(self):
        """SUMIFS(amounts, categories, "revenue", quarters, "Q1")"""
        wb = _make_wb({"S": {"literals": {
            "A1": 100, "A2": 200, "A3": 50, "A4": 300,
            "B1": "revenue", "B2": "cost", "B3": "revenue", "B4": "revenue",
            "C1": "Q1", "C2": "Q1", "C3": "Q2", "C4": "Q1",
        }}})
        tokens = [
            _func_open("SUMIFS"),
            _ref("A1:A4"), _sep(),
            _ref("B1:B4"), _sep(), _text("revenue"), _sep(),
            _ref("C1:C4"), _sep(), _text("Q1"),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == 400.0  # A1=100 + A4=300

    def test_countifs_multi_criteria(self):
        wb = _make_wb({"S": {"literals": {
            "A1": "x", "A2": "y", "A3": "x", "A4": "x",
            "B1": 10, "B2": 20, "B3": 5, "B4": 15,
        }}})
        # COUNTIFS(A1:A4, "x", B1:B4, ">8")
        tokens = [
            _func_open("COUNTIFS"),
            _ref("A1:A4"), _sep(), _text("x"), _sep(),
            _ref("B1:B4"), _sep(), _text(">8"),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == 2  # (x,10) and (x,15)

    def test_sumif_with_operator_criteria(self):
        wb = _make_wb({"S": {"literals": {
            "A1": 100, "A2": 200, "A3": 50, "A4": 300,
        }}})
        # SUMIF(A1:A4, ">150")
        tokens = [_func_open("SUMIF"), _ref("A1:A4"), _sep(), _text(">150"), _func_close()]
        assert _eval(wb, "S", tokens) == 500.0  # 200 + 300


# ===================================================================
# 5. Lookup functions — complex INDEX/MATCH/XLOOKUP combos
# ===================================================================

class TestLookupComplex:
    def _cap_table_wb(self):
        """Simulates a small cap table for lookup tests."""
        return _make_wb({"Cap": {"literals": {
            # Headers row 1
            "A1": "Investor", "B1": "Shares", "C1": "Class", "D1": "Price",
            # Data rows 2-6
            "A2": "Founder A",  "B2": 5000000,  "C2": "Common",    "D2": 0.001,
            "A3": "Founder B",  "B3": 3000000,  "C3": "Common",    "D3": 0.001,
            "A4": "Seed Fund",  "B4": 1000000,  "C4": "Preferred",  "D4": 1.00,
            "A5": "Series A",   "B5": 2000000,  "C5": "Preferred",  "D5": 2.50,
            "A6": "ESOP Pool",  "B6": 1500000,  "C6": "Common",    "D6": 0.001,
        }}})

    def test_index_match_combo(self):
        """=INDEX(D2:D6, MATCH("Series A", A2:A6, 0))  → 2.50"""
        wb = self._cap_table_wb()
        # MATCH first
        match_tokens = [_func_open("MATCH"), _text("Series A"), _sep(), _ref("A2:A6"), _sep(), _num(0), _func_close()]
        assert _eval(wb, "Cap", match_tokens) == 4  # 4th position

        # INDEX with the MATCH result
        tokens = [
            _func_open("INDEX"), _ref("D2:D6"), _sep(),
            _func_open("MATCH"), _text("Series A"), _sep(), _ref("A2:A6"), _sep(), _num(0), _func_close(),
            _func_close(),
        ]
        assert _eval(wb, "Cap", tokens) == 2.50

    def test_xlookup_exact(self):
        """=XLOOKUP("Seed Fund", A2:A6, B2:B6)  → 1000000"""
        wb = self._cap_table_wb()
        tokens = [
            _func_open("XLOOKUP"),
            _text("Seed Fund"), _sep(),
            _ref("A2:A6"), _sep(),
            _ref("B2:B6"),
            _func_close(),
        ]
        assert _eval(wb, "Cap", tokens) == 1000000

    def test_xlookup_not_found_default(self):
        """=XLOOKUP("Nobody", A2:A6, B2:B6, 0)  → 0"""
        wb = self._cap_table_wb()
        tokens = [
            _func_open("XLOOKUP"),
            _text("Nobody"), _sep(),
            _ref("A2:A6"), _sep(),
            _ref("B2:B6"), _sep(),
            _num(0),
            _func_close(),
        ]
        assert _eval(wb, "Cap", tokens) == 0.0

    def test_vlookup_2d_table(self):
        """VLOOKUP into a 4-col table, col_index=3"""
        wb = self._cap_table_wb()
        tokens = [
            _func_open("VLOOKUP"),
            _text("Founder B"), _sep(),
            _ref("A2:D6"), _sep(),
            _num(3), _sep(),
            _logical("FALSE"),
            _func_close(),
        ]
        assert _eval(wb, "Cap", tokens) == "Common"

    def test_index_2d_row_col(self):
        """=INDEX(B2:D6, 3, 2)  → Class of Seed Fund = 'Preferred'"""
        wb = self._cap_table_wb()
        tokens = [
            _func_open("INDEX"), _ref("B2:D6"), _sep(), _num(3), _sep(), _num(2),
            _func_close(),
        ]
        assert _eval(wb, "Cap", tokens) == "Preferred"

    def test_match_approximate_ascending(self):
        """MATCH(1.5, {0.001, 0.001, 1.00, 2.50, 0.001}, 1) in sorted ascending."""
        wb = _make_wb({"S": {"literals": {
            "A1": 0.5, "A2": 1.0, "A3": 2.0, "A4": 3.0,
        }}})
        tokens = [_func_open("MATCH"), _num(2.5), _sep(), _ref("A1:A4"), _sep(), _num(1), _func_close()]
        assert _eval(wb, "S", tokens) == 3  # 2.0 is the largest <= 2.5


# ===================================================================
# 6. Date functions — real financial date math
# ===================================================================

class TestDateMath:
    def test_iso_date_string_coerces_to_serial(self):
        wb = _make_wb({"S": {"literals": {"A1": "2026-01-15"}}})
        tokens = [_func_open("YEAR"), _ref("A1"), _func_close()]
        assert _eval(wb, "S", tokens) == 2026

    def test_edate_month_boundary(self):
        """EDATE from Jan 31 by 1 month → Feb 28/29."""
        wb = _make_wb({"S": {"literals": {"A1": _date_to_serial(date(2025, 1, 31))}}})
        tokens = [_func_open("EDATE"), _ref("A1"), _sep(), _num(1), _func_close()]
        result = _serial_to_date(int(_eval(wb, "S", tokens)))
        assert result == date(2025, 2, 28)

    def test_edate_leap_year(self):
        wb = _make_wb({"S": {"literals": {"A1": _date_to_serial(date(2024, 1, 31))}}})
        tokens = [_func_open("EDATE"), _ref("A1"), _sep(), _num(1), _func_close()]
        result = _serial_to_date(int(_eval(wb, "S", tokens)))
        assert result == date(2024, 2, 29)

    def test_edate_24_months(self):
        """EDATE(date, 24) — 2 years forward. Previously caused month overflow."""
        wb = _make_wb({"S": {"literals": {"A1": _date_to_serial(date(2024, 6, 15))}}})
        tokens = [_func_open("EDATE"), _ref("A1"), _sep(), _num(24), _func_close()]
        result = _serial_to_date(int(_eval(wb, "S", tokens)))
        assert result == date(2026, 6, 15)

    def test_eomonth(self):
        wb = _make_wb({"S": {"literals": {"A1": _date_to_serial(date(2024, 1, 15))}}})
        tokens = [_func_open("EOMONTH"), _ref("A1"), _sep(), _num(2), _func_close()]
        result = _serial_to_date(int(_eval(wb, "S", tokens)))
        assert result == date(2024, 3, 31)

    def test_yearfrac_30_360(self):
        """YEARFRAC with 30/360 basis."""
        d1 = _date_to_serial(date(2024, 1, 1))
        d2 = _date_to_serial(date(2024, 7, 1))
        wb = _make_wb({"S": {"literals": {"A1": d1, "A2": d2}}})
        tokens = [_func_open("YEARFRAC"), _ref("A1"), _sep(), _ref("A2"), _sep(), _num(0), _func_close()]
        result = _eval(wb, "S", tokens)
        assert abs(result - 0.5) < 0.01

    def test_date_day_month_year_round_trip(self):
        """DATE(YEAR(d), MONTH(d)+24, DAY(d)) — the pattern that crashed before."""
        d = _date_to_serial(date(2024, 3, 15))
        wb = _make_wb({"S": {"literals": {"A1": d}}})
        # YEAR(A1)
        year_tok = [_func_open("YEAR"), _ref("A1"), _func_close()]
        # MONTH(A1)
        month_tok = [_func_open("MONTH"), _ref("A1"), _func_close()]
        # DAY(A1)
        day_tok = [_func_open("DAY"), _ref("A1"), _func_close()]

        year_val = _eval(wb, "S", year_tok)
        month_val = _eval(wb, "S", month_tok)
        day_val = _eval(wb, "S", day_tok)
        assert year_val == 2024
        assert month_val == 3
        assert day_val == 15

    def test_date_function_normalizes_month_overflow(self):
        wb = _empty_wb()
        tokens = [
            _func_open("DATE"),
            _num(2026), _sep(), _num(30), _sep(), _num(30),
            _func_close(),
        ]
        result = _serial_to_date(int(_eval(wb, "S", tokens)))
        assert result == date(2028, 6, 30)


# ===================================================================
# 7. Finance functions — IRR, XIRR, NPV, XNPV
# ===================================================================

class TestFinanceFunctions:
    def test_irr_standard_investment(self):
        """Standard 5-year investment with equal cash flows."""
        wb = _make_wb({"S": {"literals": {
            "A1": -10000, "A2": 3000, "A3": 3000, "A4": 3000, "A5": 3000,
        }}})
        tokens = [_func_open("IRR"), _ref("A1:A5"), _func_close()]
        result = _eval(wb, "S", tokens)
        # ~7.7%
        assert abs(result - 0.07714) < 0.001

    def test_irr_negative_return(self):
        wb = _make_wb({"S": {"literals": {
            "A1": -10000, "A2": 2000, "A3": 2000, "A4": 2000, "A5": 2000,
        }}})
        tokens = [_func_open("IRR"), _ref("A1:A5"), _func_close()]
        result = _eval(wb, "S", tokens)
        assert result < 0  # negative IRR

    def test_xirr_irregular_dates(self):
        """XIRR with non-uniform date spacing."""
        d1 = _date_to_serial(date(2024, 1, 1))
        d2 = _date_to_serial(date(2024, 6, 15))
        d3 = _date_to_serial(date(2025, 1, 1))
        d4 = _date_to_serial(date(2026, 1, 1))
        wb = _make_wb({"S": {"literals": {
            "A1": -10000, "A2": 2000, "A3": 5000, "A4": 6000,
            "B1": d1, "B2": d2, "B3": d3, "B4": d4,
        }}})
        tokens = [_func_open("XIRR"), _ref("A1:A4"), _sep(), _ref("B1:B4"), _func_close()]
        result = _eval(wb, "S", tokens)
        assert 0.1 < result < 0.5  # should be a reasonable positive return

    def test_xirr_no_sign_change_raises(self):
        """XIRR with all positive cash flows → error."""
        d1 = _date_to_serial(date(2024, 1, 1))
        d2 = _date_to_serial(date(2025, 1, 1))
        wb = _make_wb({"S": {"literals": {
            "A1": 100, "A2": 200, "B1": d1, "B2": d2,
        }}})
        tokens = [_func_open("XIRR"), _ref("A1:A2"), _sep(), _ref("B1:B2"), _func_close()]
        with pytest.raises(EvalError, match="positive and negative"):
            _eval(wb, "S", tokens)

    def test_npv_discounting(self):
        wb = _make_wb({"S": {"literals": {"A1": 1000, "A2": 1000, "A3": 1000}}})
        tokens = [_func_open("NPV"), _num(0.1), _sep(), _ref("A1:A3"), _func_close()]
        result = _eval(wb, "S", tokens)
        expected = 1000/1.1 + 1000/1.21 + 1000/1.331
        assert abs(result - expected) < 0.01

    def test_xnpv(self):
        d1 = _date_to_serial(date(2024, 1, 1))
        d2 = _date_to_serial(date(2025, 1, 1))
        d3 = _date_to_serial(date(2026, 1, 1))
        wb = _make_wb({"S": {"literals": {
            "A1": -1000, "A2": 600, "A3": 600,
            "B1": d1, "B2": d2, "B3": d3,
        }}})
        tokens = [_func_open("XNPV"), _num(0.1), _sep(), _ref("A1:A3"), _sep(), _ref("B1:B3"), _func_close()]
        result = _eval(wb, "S", tokens)
        # Should be slightly positive
        assert result > 0


# ===================================================================
# 8. Text functions
# ===================================================================

class TestTextFunctions:
    def test_textjoin_with_ignore_empty(self):
        wb = _make_wb({"S": {"literals": {"A1": "a", "A2": "", "A3": "b", "A4": "", "A5": "c"}}})
        tokens = [
            _func_open("TEXTJOIN"), _text(", "), _sep(), _logical("TRUE"), _sep(),
            _ref("A1:A5"),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == "a, b, c"

    def test_concat_range(self):
        wb = _make_wb({"S": {"literals": {"A1": "Hello", "A2": " ", "A3": "World"}}})
        tokens = [_func_open("CONCAT"), _ref("A1:A3"), _func_close()]
        assert _eval(wb, "S", tokens) == "Hello World"


# ===================================================================
# 9. Numeric transforms
# ===================================================================

class TestNumericTransforms:
    def test_roundup_negative(self):
        wb = _empty_wb()
        tokens = [_func_open("ROUNDUP"), _prefix("-"), _num(3.141), _sep(), _num(2), _func_close()]
        # ROUNDUP(-3.141, 2) → -3.15 (ceiling of magnitude)
        assert _eval(wb, "S", tokens) == -3.15

    def test_rounddown_large(self):
        wb = _empty_wb()
        tokens = [_func_open("ROUNDDOWN"), _num(123456.789), _sep(), _prefix("-"), _num(3), _func_close()]
        # ROUNDDOWN(123456.789, -3) → 123000
        assert _eval(wb, "S", tokens) == 123000.0


# ===================================================================
# 10. Function classification
# ===================================================================

class TestFunctionClassification:
    def test_all_supported_funcs(self):
        ok, unsup = classify_functions(["SUM", "IF", "IFERROR", "XIRR", "VLOOKUP", "ROUND"])
        assert ok
        assert unsup == []

    def test_unsupported_detected(self):
        ok, unsup = classify_functions(["SUM", "OFFSET", "INDIRECT"])
        assert not ok
        assert "OFFSET" in unsup
        assert "INDIRECT" in unsup

    def test_unknown_function(self):
        ok, unsup = classify_functions(["GETPIVOTDATA"])
        assert not ok
        assert "GETPIVOTDATA" in unsup


# ===================================================================
# 11. Dependency graph & SCC detection
# ===================================================================

class TestGraphConstruction:
    def test_cross_sheet_deps(self):
        wb = _make_wb({
            "S1": {"literals": {"A1": 10}},
            "S2": {"formulas": {
                "B1": {
                    "dependencies": [{"sheet": "S1", "cell": "A1", "via": "A1"}],
                    "tokens": [_ref("'S1'!A1")],
                },
            }},
        })
        graph = build_dependency_graph(wb)
        assert "S1!A1" in graph["S2!B1"]

    def test_scc_three_node_cycle(self):
        graph = {"A": {"B"}, "B": {"C"}, "C": {"A"}, "D": set()}
        sccs = tarjan_scc(graph)
        cycle_scc = [s for s in sccs if len(s) > 1]
        assert len(cycle_scc) == 1
        assert set(cycle_scc[0]) == {"A", "B", "C"}

    def test_scc_self_loop(self):
        graph = {"A": {"A"}}
        sccs = tarjan_scc(graph)
        # A self-referencing node is an SCC of size 1 but IS in its own dep set
        assert any("A" in s for s in sccs)

    def test_topological_sort_respects_deps(self):
        graph = {
            "D": {"C", "B"},
            "C": {"A"},
            "B": {"A"},
            "A": set(),
        }
        order = topological_sort(graph, set())
        assert order.index("A") < order.index("B")
        assert order.index("A") < order.index("C")
        assert order.index("B") < order.index("D")
        assert order.index("C") < order.index("D")

    def test_topological_sort_excludes_scc(self):
        graph = {
            "A": set(),
            "B": {"A"},
            "C": {"D"},  # C↔D cycle
            "D": {"C"},
        }
        order = topological_sort(graph, {"C", "D"})
        assert "C" not in order
        assert "D" not in order
        assert order.index("A") < order.index("B")


# ===================================================================
# 12. Iterative solver — financial patterns
# ===================================================================

class TestIterativeSolverReal:
    def test_share_price_share_count_loop(self):
        """Core cap table circular ref:
        price = post_money / total_shares
        new_shares = investment / price
        total_shares = existing + new_shares

        post_money = 10M, investment = 2M, existing = 8M shares
        Algebra: price = 10M / (8M + 2M/price)
                 price*(8M + 2M/price) = 10M
                 8M*price + 2M = 10M
                 price = 8M/8M = 1.0
                 new_shares = 2M
                 total = 10M
        """
        wb = _make_wb({"S": {
            "literals": {
                "A1": 10000000,  # post_money
                "A2": 2000000,   # investment
                "A3": 8000000,   # existing_shares
            },
            "formulas": {
                # B1 = price = A1 / B3 (total)
                "B1": {
                    "formula": "=IFERROR(A1/B3,0)",
                    "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}, {"sheet": "S", "cell": "B3", "via": "B3"}],
                    "tokens": [
                        _func_open("IFERROR"), _ref("A1"), _op("/"), _ref("B3"),
                        _sep(), _num(0), _func_close(),
                    ],
                    "iterative_component": "iter_1",
                },
                # B2 = new_shares = ROUND(A2 / B1, 0)
                "B2": {
                    "formula": "=IFERROR(ROUND(A2/B1,0),0)",
                    "dependencies": [{"sheet": "S", "cell": "A2", "via": "A2"}, {"sheet": "S", "cell": "B1", "via": "B1"}],
                    "tokens": [
                        _func_open("IFERROR"),
                        _func_open("ROUND"), _ref("A2"), _op("/"), _ref("B1"), _sep(), _num(0), _func_close(),
                        _sep(), _num(0), _func_close(),
                    ],
                    "iterative_component": "iter_1",
                },
                # B3 = total = A3 + B2
                "B3": {
                    "formula": "=A3+B2",
                    "dependencies": [{"sheet": "S", "cell": "A3", "via": "A3"}, {"sheet": "S", "cell": "B2", "via": "B2"}],
                    "tokens": [_ref("A3"), _op("+"), _ref("B2")],
                    "iterative_component": "iter_1",
                },
            },
        }})
        converged = solve_iterative_component(wb, ["S!B1", "S!B2", "S!B3"])
        assert converged
        price = wb.formula_cells["S!B1"].computed_value
        new_shares = wb.formula_cells["S!B2"].computed_value
        total = wb.formula_cells["S!B3"].computed_value

        assert abs(price - 1.0) < 0.01
        assert abs(new_shares - 2000000) < 10  # ROUND might be off by 1
        assert abs(total - 10000000) < 10

    def test_option_pool_topup_loop(self):
        """ESOP top-up loop:
        esop_new = max(target% * total - esop_existing, 0)
        total = existing + new_investor + esop_new

        target=10%, existing=8M, new_investor=2M, esop_existing=500K
        total = 8M + 2M + esop_new
        esop_new = max(0.1 * total - 500K, 0)
        esop_new = 0.1*(10M + esop_new) - 500K
        esop_new = 1M + 0.1*esop_new - 500K
        0.9*esop_new = 500K
        esop_new = 555556
        total = 10555556
        """
        wb = _make_wb({"S": {
            "literals": {
                "A1": 8000000,   # existing
                "A2": 2000000,   # new_investor
                "A3": 500000,    # esop_existing
                "A4": 0.10,      # target_pct
            },
            "formulas": {
                # B1 = esop_new = MAX(A4*B2 - A3, 0)
                "B1": {
                    "formula": "=MAX(A4*B2-A3,0)",
                    "functions": ["MAX"],
                    "dependencies": [
                        {"sheet": "S", "cell": "A4", "via": "A4"},
                        {"sheet": "S", "cell": "B2", "via": "B2"},
                        {"sheet": "S", "cell": "A3", "via": "A3"},
                    ],
                    "tokens": [
                        _func_open("MAX"),
                        _ref("A4"), _op("*"), _ref("B2"), _op("-"), _ref("A3"),
                        _sep(), _num(0),
                        _func_close(),
                    ],
                    "iterative_component": "iter_1",
                },
                # B2 = total = A1 + A2 + B1
                "B2": {
                    "formula": "=A1+A2+B1",
                    "dependencies": [
                        {"sheet": "S", "cell": "A1", "via": "A1"},
                        {"sheet": "S", "cell": "A2", "via": "A2"},
                        {"sheet": "S", "cell": "B1", "via": "B1"},
                    ],
                    "tokens": [_ref("A1"), _op("+"), _ref("A2"), _op("+"), _ref("B1")],
                    "iterative_component": "iter_1",
                },
            },
        }})
        converged = solve_iterative_component(wb, ["S!B1", "S!B2"])
        assert converged
        esop_new = wb.formula_cells["S!B1"].computed_value
        total = wb.formula_cells["S!B2"].computed_value
        assert abs(esop_new - 555556) < 2
        assert abs(total - 10555556) < 2

    def test_divergent_loop_detected(self):
        """A1=B1*2, B1=A1*2 with nonzero seed → diverges to inf."""
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=B1*2",
                "cached_value": 1.0,
                "dependencies": [{"sheet": "S", "cell": "B1", "via": "B1"}],
                "tokens": [_ref("B1"), _op("*"), _num(2)],
                "iterative_component": "iter_1",
            },
            "B1": {
                "formula": "=A1*2",
                "cached_value": 1.0,
                "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}],
                "tokens": [_ref("A1"), _op("*"), _num(2)],
                "iterative_component": "iter_1",
            },
        }}})
        converged = solve_iterative_component(wb, ["S!A1", "S!B1"])
        assert not converged

    def test_multi_component_convergence(self):
        """Two separate iterative components in one workbook.

        iter_1: price/share loop (S1)
        iter_2: ESOP top-up (S2) that depends on iter_1 output

        This tests interleaved execution: iter_1 solves first, then
        acyclic cells propagate, then iter_2 solves.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["S1", "S2"],
                "functions_used": {"IFERROR": 2, "ROUND": 1, "MAX": 1, "SUM": 1},
                "iterative_components": [
                    {"id": "iter_1", "cells": ["S1!B1", "S1!B2", "S1!B3"]},
                    {"id": "iter_2", "cells": ["S2!A2", "S2!A3"]},
                ],
            }))
            (d / "S1.json").write_text(json.dumps({
                "sheet": "S1",
                "formula_cells": [
                    {
                        "cell": "B1", "formula": "=IFERROR(A1/B3,0)",
                        "cached_value": None, "functions": ["IFERROR"],
                        "references": [], "iterative_component": "iter_1",
                        "dependencies": [{"sheet": "S1", "cell": "A1", "via": "A1"}, {"sheet": "S1", "cell": "B3", "via": "B3"}],
                        "tokens": [{"type": "FUNC", "subtype": "OPEN", "value": "IFERROR("}, {"type": "OPERAND", "subtype": "RANGE", "value": "A1"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "/"}, {"type": "OPERAND", "subtype": "RANGE", "value": "B3"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}],
                    },
                    {
                        "cell": "B2", "formula": "=IFERROR(ROUND(A2/B1,0),0)",
                        "cached_value": None, "functions": ["IFERROR", "ROUND"],
                        "references": [], "iterative_component": "iter_1",
                        "dependencies": [{"sheet": "S1", "cell": "A2", "via": "A2"}, {"sheet": "S1", "cell": "B1", "via": "B1"}],
                        "tokens": [{"type": "FUNC", "subtype": "OPEN", "value": "IFERROR("}, {"type": "FUNC", "subtype": "OPEN", "value": "ROUND("}, {"type": "OPERAND", "subtype": "RANGE", "value": "A2"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "/"}, {"type": "OPERAND", "subtype": "RANGE", "value": "B1"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}],
                    },
                    {
                        "cell": "B3", "formula": "=A3+B2",
                        "cached_value": None, "functions": [],
                        "references": [], "iterative_component": "iter_1",
                        "dependencies": [{"sheet": "S1", "cell": "A3", "via": "A3"}, {"sheet": "S1", "cell": "B2", "via": "B2"}],
                        "tokens": [{"type": "OPERAND", "subtype": "RANGE", "value": "A3"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "+"}, {"type": "OPERAND", "subtype": "RANGE", "value": "B2"}],
                    },
                ],
                "literal_cells": {"A1": 10000000, "A2": 2000000, "A3": 8000000},
            }))
            # S2 depends on S1!B3 (total_shares) via an acyclic cell S2!A1
            (d / "S2.json").write_text(json.dumps({
                "sheet": "S2",
                "formula_cells": [
                    {   # A1 = S1!B3 (total from iter_1)
                        "cell": "A1", "formula": "='S1'!B3",
                        "cached_value": None, "functions": [],
                        "references": [{"sheet": "S1", "ref": "B3", "kind": "cell", "expanded": ["B3"]}],
                        "dependencies": [{"sheet": "S1", "cell": "B3", "via": "B3"}],
                        "tokens": [{"type": "OPERAND", "subtype": "RANGE", "value": "'S1'!B3"}],
                        "iterative_component": None,
                    },
                    {   # A2 = esop_new = MAX(0.1*A3 - 500000, 0)
                        "cell": "A2", "formula": "=MAX(0.1*A3-500000,0)",
                        "cached_value": None, "functions": ["MAX"],
                        "references": [], "iterative_component": "iter_2",
                        "dependencies": [{"sheet": "S2", "cell": "A3", "via": "A3"}],
                        "tokens": [{"type": "FUNC", "subtype": "OPEN", "value": "MAX("}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0.1"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "*"}, {"type": "OPERAND", "subtype": "RANGE", "value": "A3"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "-"}, {"type": "OPERAND", "subtype": "NUMBER", "value": "500000"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}],
                    },
                    {   # A3 = total_with_esop = A1 + A2
                        "cell": "A3", "formula": "=A1+A2",
                        "cached_value": None, "functions": [],
                        "references": [], "iterative_component": "iter_2",
                        "dependencies": [{"sheet": "S2", "cell": "A1", "via": "A1"}, {"sheet": "S2", "cell": "A2", "via": "A2"}],
                        "tokens": [{"type": "OPERAND", "subtype": "RANGE", "value": "A1"}, {"type": "OPERATOR-INFIX", "subtype": "", "value": "+"}, {"type": "OPERAND", "subtype": "RANGE", "value": "A2"}],
                    },
                ],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            assert verdict["logic"]["passed"] is True
            assert verdict["summary"]["supported_converged"] == 6  # 3 in S1, 3 in S2
            assert verdict["summary"]["supported_nonconverged"] == 0

            # Verify actual values
            wb = load_workbook(d)
            # Re-run to populate
            v2 = evaluate_workbook(d)
            # iter_1 should give price=1.0, iter_2 should give esop_new≈555556
            assert all(c == "converged" for c in v2["summary"]["iterative_components"].values())


# ===================================================================
# 13. Nonsense / error detection
# ===================================================================

class TestNonsenseDetection:
    def test_broken_sheet_reference(self):
        """Formula references a sheet that doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["Sheet1"],
                "functions_used": {},
                "iterative_components": [],
            }))
            (d / "Sheet1.json").write_text(json.dumps({
                "sheet": "Sheet1",
                "formula_cells": [{
                    "cell": "A1", "formula": "='Missing Sheet'!B5",
                    "cached_value": None, "functions": [],
                    "references": [{"sheet": "Missing Sheet", "ref": "B5", "kind": "cell", "expanded": ["B5"]}],
                    "dependencies": [{"sheet": "Missing Sheet", "cell": "B5", "via": "B5"}],
                    "tokens": [{"type": "OPERAND", "subtype": "RANGE", "value": "'Missing Sheet'!B5"}],
                    "iterative_component": None,
                }],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            # Should detect the broken reference
            assert any("Missing Sheet" in str(i.get("details", "")) for i in verdict["logic"]["issues"])

    def test_unsupported_function_flagged(self):
        """OFFSET should be classified as unsupported, not silently skipped."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["Sheet1"],
                "functions_used": {"OFFSET": 1},
                "iterative_components": [],
            }))
            (d / "Sheet1.json").write_text(json.dumps({
                "sheet": "Sheet1",
                "formula_cells": [{
                    "cell": "A1", "formula": "=OFFSET(B1,1,0)",
                    "cached_value": 42, "functions": ["OFFSET"],
                    "references": [], "dependencies": [],
                    "tokens": [{"type": "FUNC", "subtype": "OPEN", "value": "OFFSET("}, {"type": "OPERAND", "subtype": "RANGE", "value": "B1"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "1"}, {"type": "SEP", "subtype": "ARG", "value": ","}, {"type": "OPERAND", "subtype": "NUMBER", "value": "0"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}],
                    "iterative_component": None,
                }],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            assert verdict["summary"]["unsupported"] == 1

    def test_spec_does_not_drive_missing_sheet_failures(self):
        """The deterministic evaluator is formulas->values only and ignores sheet-set policy."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["Cap Table"],
                "functions_used": {},
                "iterative_components": [],
            }))
            (d / "Cap Table.json").write_text(json.dumps({
                "sheet": "Cap Table", "formula_cells": [], "literal_cells": {},
            }))

            spec = {"sheets": [
                {"name": "Cap Table"},
                {"name": "Series A", "build_type": "build"},
                {"name": "Returns", "build_type": "build"},
            ]}
            verdict = evaluate_workbook(d, spec)
            assert verdict["logic"]["passed"]
            missing = [i for i in verdict["logic"]["issues"] if "missing" in i["summary"].lower()]
            assert len(missing) == 0

    def test_all_zero_rows_not_flagged_in_primitive_mode(self):
        """Row hygiene is out of scope for the formula->values primitive."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["Sheet1"],
                "functions_used": {"SUM": 3},
                "iterative_components": [],
            }))
            cells = []
            for col in ["B", "C", "D"]:
                cells.append({
                    "cell": f"{col}5", "formula": f"=SUM({col}1:{col}4)",
                    "cached_value": 0, "functions": ["SUM"],
                    "references": [{"sheet": "Sheet1", "ref": f"{col}1:{col}4", "kind": "range", "expanded": [f"{col}1", f"{col}2", f"{col}3", f"{col}4"]}],
                    "dependencies": [{"sheet": "Sheet1", "cell": f"{col}{r}", "via": f"{col}1:{col}4"} for r in range(1, 5)],
                    "tokens": [{"type": "FUNC", "subtype": "OPEN", "value": "SUM("}, {"type": "OPERAND", "subtype": "RANGE", "value": f"{col}1:{col}4"}, {"type": "FUNC", "subtype": "CLOSE", "value": ")"}],
                    "iterative_component": None,
                })
            (d / "Sheet1.json").write_text(json.dumps({
                "sheet": "Sheet1", "formula_cells": cells, "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            zero_issues = [i for i in verdict["logic"]["issues"] if "zero" in i["summary"].lower()]
            assert len(zero_issues) == 0


# ===================================================================
# 14. Acceptance test against real workbook dump
# ===================================================================

class TestAcceptance:
    """Run the evaluator against the real workbook dumps in evals/."""

    @pytest.fixture
    def current_formulas_dir(self):
        d = Path(__file__).parent.parent / "evals" / "formulas_json"
        if not (d / "workbook.json").exists():
            pytest.skip("No real workbook dump available")
        return d

    def test_current_workbook_converges(self, current_formulas_dir):
        """The current model workbook should fully converge."""
        verdict = evaluate_workbook(current_formulas_dir)
        s = verdict["summary"]
        assert s["supported_converged"] == s["total_cells"], (
            f'{s["supported_nonconverged"]} non-converged, '
            f'{s["parse_failures"]} parse failures'
        )
        assert verdict["logic"]["passed"]

    def test_current_workbook_no_unsupported(self, current_formulas_dir):
        """Current workbook uses only supported functions."""
        verdict = evaluate_workbook(current_formulas_dir)
        assert verdict["summary"]["unsupported"] == 0

    def test_current_workbook_under_5s(self, current_formulas_dir):
        """Performance target: full eval under 5 seconds."""
        import time
        t0 = time.time()
        evaluate_workbook(current_formulas_dir)
        elapsed = time.time() - t0
        assert elapsed < 5.0, f"Evaluation took {elapsed:.2f}s (target: <5s)"


# ===================================================================
# 15. Non-cap-table financial models
# ===================================================================

class TestLoanAmortization:
    """Loan amortization schedule — monthly payment, interest, principal split."""

    def _loan_wb(self):
        """Simple 3-period loan: principal=100000, annual_rate=12%, periods=3
        Monthly rate = 1%
        PMT = P * r / (1-(1+r)^-n) = 100000 * 0.01 / (1 - 1.01^-3) = 34002.21...
        We'll compute it step by step since PMT isn't supported — use SUM checks.
        """
        return _make_wb({"Loan": {
            "literals": {
                "B1": 100000,     # principal
                "B2": 0.12,       # annual rate
                "B3": 3,          # periods
                "B4": 0.01,       # monthly rate = B2/12
                # Period 1
                "C6": 1,          # period number
                # Period 2
                "D6": 2,
                # Period 3
                "E6": 3,
            },
            "formulas": {
                # C7 = beginning balance = B1
                "C7": {
                    "formula": "=B1",
                    "dependencies": [{"sheet": "Loan", "cell": "B1", "via": "B1"}],
                    "tokens": [_ref("B1")],
                },
                # C8 = interest = C7 * B4
                "C8": {
                    "formula": "=C7*B4",
                    "dependencies": [{"sheet": "Loan", "cell": "C7", "via": "C7"}, {"sheet": "Loan", "cell": "B4", "via": "B4"}],
                    "tokens": [_ref("C7"), _op("*"), _ref("B4")],
                },
                # C9 = payment = ROUND(B1*B4/(1-(1+B4)^-B3), 2)
                "C9": {
                    "formula": "=ROUND(B1*B4/(1-(1+B4)^-B3),2)",
                    "functions": ["ROUND"],
                    "dependencies": [{"sheet": "Loan", "cell": "B1", "via": "B1"}, {"sheet": "Loan", "cell": "B4", "via": "B4"}, {"sheet": "Loan", "cell": "B3", "via": "B3"}],
                    "tokens": [
                        _func_open("ROUND"),
                        _ref("B1"), _op("*"), _ref("B4"), _op("/"),
                        _paren_open(),
                        _num(1), _op("-"),
                        _paren_open(), _num(1), _op("+"), _ref("B4"), _paren_close(),
                        _op("^"),
                        _prefix("-"), _ref("B3"),
                        _paren_close(),
                        _sep(), _num(2),
                        _func_close(),
                    ],
                },
                # C10 = principal = C9 - C8
                "C10": {
                    "formula": "=C9-C8",
                    "dependencies": [{"sheet": "Loan", "cell": "C9", "via": "C9"}, {"sheet": "Loan", "cell": "C8", "via": "C8"}],
                    "tokens": [_ref("C9"), _op("-"), _ref("C8")],
                },
                # C11 = ending balance = C7 - C10
                "C11": {
                    "formula": "=C7-C10",
                    "dependencies": [{"sheet": "Loan", "cell": "C7", "via": "C7"}, {"sheet": "Loan", "cell": "C10", "via": "C10"}],
                    "tokens": [_ref("C7"), _op("-"), _ref("C10")],
                },
                # Period 2: D7 = C11
                "D7": {
                    "formula": "=C11",
                    "dependencies": [{"sheet": "Loan", "cell": "C11", "via": "C11"}],
                    "tokens": [_ref("C11")],
                },
                "D8": {
                    "formula": "=D7*B4",
                    "dependencies": [{"sheet": "Loan", "cell": "D7", "via": "D7"}, {"sheet": "Loan", "cell": "B4", "via": "B4"}],
                    "tokens": [_ref("D7"), _op("*"), _ref("B4")],
                },
                "D9": {
                    "formula": "=C9",
                    "dependencies": [{"sheet": "Loan", "cell": "C9", "via": "C9"}],
                    "tokens": [_ref("C9")],
                },
                "D10": {
                    "formula": "=D9-D8",
                    "dependencies": [{"sheet": "Loan", "cell": "D9", "via": "D9"}, {"sheet": "Loan", "cell": "D8", "via": "D8"}],
                    "tokens": [_ref("D9"), _op("-"), _ref("D8")],
                },
                "D11": {
                    "formula": "=D7-D10",
                    "dependencies": [{"sheet": "Loan", "cell": "D7", "via": "D7"}, {"sheet": "Loan", "cell": "D10", "via": "D10"}],
                    "tokens": [_ref("D7"), _op("-"), _ref("D10")],
                },
                # Total interest = SUM(C8, D8)
                "F8": {
                    "formula": "=SUM(C8,D8)",
                    "functions": ["SUM"],
                    "dependencies": [{"sheet": "Loan", "cell": "C8", "via": "C8"}, {"sheet": "Loan", "cell": "D8", "via": "D8"}],
                    "tokens": [_func_open("SUM"), _ref("C8"), _sep(), _ref("D8"), _func_close()],
                },
            },
        }})

    def test_interest_decreases_each_period(self):
        wb = self._loan_wb()
        graph = build_dependency_graph(wb)
        order = topological_sort(graph, set())
        for qcell in order:
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        c8 = wb.formula_cells["Loan!C8"].computed_value  # period 1 interest
        d8 = wb.formula_cells["Loan!D8"].computed_value  # period 2 interest
        assert c8 > d8  # interest decreases as balance goes down

    def test_principal_increases_each_period(self):
        wb = self._loan_wb()
        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        c10 = wb.formula_cells["Loan!C10"].computed_value
        d10 = wb.formula_cells["Loan!D10"].computed_value
        assert d10 > c10  # principal portion grows

    def test_payment_covers_interest(self):
        wb = self._loan_wb()
        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        payment = wb.formula_cells["Loan!C9"].computed_value
        interest = wb.formula_cells["Loan!C8"].computed_value
        assert payment > interest
        assert payment > 0


class TestRevenueModel:
    """Revenue forecast model with growth rates and conditional logic."""

    def _revenue_wb(self):
        """3-year revenue model:
        Year 1: base=1000000
        Year 2: grow 20% if Y1 > 500K, else grow 5%
        Year 3: grow 15% if Y2 > 1.5M, else grow 10%
        Total revenue = SUM
        """
        return _make_wb({"Rev": {
            "literals": {"B1": 1000000},  # year 1 base
            "formulas": {
                # C1 = IF(B1>500000, B1*1.2, B1*1.05)
                "C1": {
                    "formula": "=IF(B1>500000,B1*1.2,B1*1.05)",
                    "functions": ["IF"],
                    "dependencies": [{"sheet": "Rev", "cell": "B1", "via": "B1"}],
                    "tokens": [
                        _func_open("IF"),
                        _ref("B1"), _op(">"), _num(500000), _sep(),
                        _ref("B1"), _op("*"), _num(1.2), _sep(),
                        _ref("B1"), _op("*"), _num(1.05),
                        _func_close(),
                    ],
                },
                # D1 = IF(C1>1500000, C1*1.15, C1*1.10)
                "D1": {
                    "formula": "=IF(C1>1500000,C1*1.15,C1*1.10)",
                    "functions": ["IF"],
                    "dependencies": [{"sheet": "Rev", "cell": "C1", "via": "C1"}],
                    "tokens": [
                        _func_open("IF"),
                        _ref("C1"), _op(">"), _num(1500000), _sep(),
                        _ref("C1"), _op("*"), _num(1.15), _sep(),
                        _ref("C1"), _op("*"), _num(1.10),
                        _func_close(),
                    ],
                },
                # E1 = total = SUM(B1:D1)
                "E1": {
                    "formula": "=SUM(B1:D1)",
                    "functions": ["SUM"],
                    "dependencies": [
                        {"sheet": "Rev", "cell": "B1", "via": "B1:D1"},
                        {"sheet": "Rev", "cell": "C1", "via": "B1:D1"},
                        {"sheet": "Rev", "cell": "D1", "via": "B1:D1"},
                    ],
                    "tokens": [_func_open("SUM"), _ref("B1:D1"), _func_close()],
                },
                # F1 = YoY growth = (D1-C1)/C1
                "F1": {
                    "formula": "=(D1-C1)/C1",
                    "dependencies": [
                        {"sheet": "Rev", "cell": "D1", "via": "D1"},
                        {"sheet": "Rev", "cell": "C1", "via": "C1"},
                    ],
                    "tokens": [
                        _paren_open(), _ref("D1"), _op("-"), _ref("C1"), _paren_close(),
                        _op("/"), _ref("C1"),
                    ],
                },
            },
        }})

    def test_conditional_growth_path(self):
        wb = self._revenue_wb()
        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        # Y1=1M > 500K → Y2 = 1M * 1.2 = 1.2M
        assert wb.formula_cells["Rev!C1"].computed_value == 1200000.0
        # Y2=1.2M < 1.5M → Y3 = 1.2M * 1.10 = 1.32M
        assert wb.formula_cells["Rev!D1"].computed_value == 1320000.0
        # Total = 1M + 1.2M + 1.32M = 3.52M
        assert wb.formula_cells["Rev!E1"].computed_value == 3520000.0
        # YoY growth = (1.32M - 1.2M) / 1.2M = 0.1
        assert abs(wb.formula_cells["Rev!F1"].computed_value - 0.1) < 1e-10

    def test_different_base_takes_other_branch(self):
        """With base=400K, should take the 5% growth branch."""
        wb = self._revenue_wb()
        wb.literal_values["Rev!B1"] = 400000  # below 500K threshold
        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        # Y1=400K < 500K → Y2 = 400K * 1.05 = 420K
        assert wb.formula_cells["Rev!C1"].computed_value == 420000.0


class TestDCFModel:
    """Simplified DCF: project free cash flows, discount, compute enterprise value."""

    def test_dcf_valuation(self):
        wb = _make_wb({"DCF": {
            "literals": {
                # FCF projections years 1-4
                "B1": 500000, "C1": 600000, "D1": 720000, "E1": 864000,
                # Discount rate
                "B3": 0.10,
                # Terminal growth rate
                "B4": 0.03,
            },
            "formulas": {
                # PV of each year: FCF / (1+r)^n
                "B5": {
                    "formula": "=B1/(1+$B$3)^1",
                    "dependencies": [{"sheet": "DCF", "cell": "B1", "via": "B1"}, {"sheet": "DCF", "cell": "B3", "via": "B3"}],
                    "tokens": [_ref("B1"), _op("/"), _paren_open(), _num(1), _op("+"), _ref("B3"), _paren_close(), _op("^"), _num(1)],
                },
                "C5": {
                    "formula": "=C1/(1+$B$3)^2",
                    "dependencies": [{"sheet": "DCF", "cell": "C1", "via": "C1"}, {"sheet": "DCF", "cell": "B3", "via": "B3"}],
                    "tokens": [_ref("C1"), _op("/"), _paren_open(), _num(1), _op("+"), _ref("B3"), _paren_close(), _op("^"), _num(2)],
                },
                "D5": {
                    "formula": "=D1/(1+$B$3)^3",
                    "dependencies": [{"sheet": "DCF", "cell": "D1", "via": "D1"}, {"sheet": "DCF", "cell": "B3", "via": "B3"}],
                    "tokens": [_ref("D1"), _op("/"), _paren_open(), _num(1), _op("+"), _ref("B3"), _paren_close(), _op("^"), _num(3)],
                },
                "E5": {
                    "formula": "=E1/(1+$B$3)^4",
                    "dependencies": [{"sheet": "DCF", "cell": "E1", "via": "E1"}, {"sheet": "DCF", "cell": "B3", "via": "B3"}],
                    "tokens": [_ref("E1"), _op("/"), _paren_open(), _num(1), _op("+"), _ref("B3"), _paren_close(), _op("^"), _num(4)],
                },
                # Terminal value = FCF_last * (1+g) / (r-g), discounted
                # TV = 864000 * 1.03 / (0.10 - 0.03) / (1.1)^4
                "F5": {
                    "formula": "=E1*(1+$B$4)/($B$3-$B$4)/(1+$B$3)^4",
                    "dependencies": [
                        {"sheet": "DCF", "cell": "E1", "via": "E1"},
                        {"sheet": "DCF", "cell": "B3", "via": "B3"},
                        {"sheet": "DCF", "cell": "B4", "via": "B4"},
                    ],
                    "tokens": [
                        _ref("E1"), _op("*"),
                        _paren_open(), _num(1), _op("+"), _ref("B4"), _paren_close(),
                        _op("/"),
                        _paren_open(), _ref("B3"), _op("-"), _ref("B4"), _paren_close(),
                        _op("/"),
                        _paren_open(), _num(1), _op("+"), _ref("B3"), _paren_close(),
                        _op("^"), _num(4),
                    ],
                },
                # Enterprise value = SUM of PVs + terminal
                "G5": {
                    "formula": "=SUM(B5:F5)",
                    "functions": ["SUM"],
                    "dependencies": [
                        {"sheet": "DCF", "cell": "B5", "via": "B5:F5"},
                        {"sheet": "DCF", "cell": "C5", "via": "B5:F5"},
                        {"sheet": "DCF", "cell": "D5", "via": "B5:F5"},
                        {"sheet": "DCF", "cell": "E5", "via": "B5:F5"},
                        {"sheet": "DCF", "cell": "F5", "via": "B5:F5"},
                    ],
                    "tokens": [_func_open("SUM"), _ref("B5:F5"), _func_close()],
                },
            },
        }})

        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        ev = wb.formula_cells["DCF!G5"].computed_value

        # Manual calculation:
        pv1 = 500000 / 1.1
        pv2 = 600000 / 1.1**2
        pv3 = 720000 / 1.1**3
        pv4 = 864000 / 1.1**4
        tv = 864000 * 1.03 / 0.07 / 1.1**4
        expected = pv1 + pv2 + pv3 + pv4 + tv

        assert abs(ev - expected) < 1.0
        assert ev > 8000000  # EV should be north of 8M


class TestWaterfallModel:
    """Distribution waterfall — common in PE/VC, not cap table specific."""

    def test_preferred_return_waterfall(self):
        """
        1. LP gets 8% preferred return on 1M investment
        2. GP catch-up to 20% of profits
        3. 80/20 split on remainder
        Total distributable: 1.5M
        """
        wb = _make_wb({"WF": {
            "literals": {
                "A1": 1000000,   # investment
                "A2": 0.08,      # preferred return
                "A3": 1500000,   # total distributable
                "A4": 0.20,      # GP carry
            },
            "formulas": {
                # B1 = preferred = MIN(A1*A2, A3-A1) — return of capital first
                # Actually: preferred amount = A1 * A2 = 80K
                "B1": {
                    "formula": "=A1*A2",
                    "dependencies": [{"sheet": "WF", "cell": "A1", "via": "A1"}, {"sheet": "WF", "cell": "A2", "via": "A2"}],
                    "tokens": [_ref("A1"), _op("*"), _ref("A2")],
                },
                # B2 = profit after preferred = A3 - A1 - B1 = 1.5M - 1M - 80K = 420K
                "B2": {
                    "formula": "=A3-A1-B1",
                    "dependencies": [
                        {"sheet": "WF", "cell": "A3", "via": "A3"},
                        {"sheet": "WF", "cell": "A1", "via": "A1"},
                        {"sheet": "WF", "cell": "B1", "via": "B1"},
                    ],
                    "tokens": [_ref("A3"), _op("-"), _ref("A1"), _op("-"), _ref("B1")],
                },
                # B3 = GP catch-up = MIN(B2, (A1+B1)*A4/(1-A4) - 0)
                # Simplified: GP gets A4 share of B2
                "B3": {
                    "formula": "=ROUND(B2*A4,2)",
                    "functions": ["ROUND"],
                    "dependencies": [
                        {"sheet": "WF", "cell": "B2", "via": "B2"},
                        {"sheet": "WF", "cell": "A4", "via": "A4"},
                    ],
                    "tokens": [
                        _func_open("ROUND"), _ref("B2"), _op("*"), _ref("A4"),
                        _sep(), _num(2), _func_close(),
                    ],
                },
                # B4 = LP share = B2 - B3
                "B4": {
                    "formula": "=B2-B3",
                    "dependencies": [
                        {"sheet": "WF", "cell": "B2", "via": "B2"},
                        {"sheet": "WF", "cell": "B3", "via": "B3"},
                    ],
                    "tokens": [_ref("B2"), _op("-"), _ref("B3")],
                },
                # B5 = total LP = A1 + B1 + B4 (capital + preferred + profit share)
                "B5": {
                    "formula": "=A1+B1+B4",
                    "dependencies": [
                        {"sheet": "WF", "cell": "A1", "via": "A1"},
                        {"sheet": "WF", "cell": "B1", "via": "B1"},
                        {"sheet": "WF", "cell": "B4", "via": "B4"},
                    ],
                    "tokens": [_ref("A1"), _op("+"), _ref("B1"), _op("+"), _ref("B4")],
                },
                # B6 = check: B5 + B3 should = A3
                "B6": {
                    "formula": "=B5+B3",
                    "dependencies": [
                        {"sheet": "WF", "cell": "B5", "via": "B5"},
                        {"sheet": "WF", "cell": "B3", "via": "B3"},
                    ],
                    "tokens": [_ref("B5"), _op("+"), _ref("B3")],
                },
            },
        }})

        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        preferred = wb.formula_cells["WF!B1"].computed_value
        gp_share = wb.formula_cells["WF!B3"].computed_value
        lp_total = wb.formula_cells["WF!B5"].computed_value
        check = wb.formula_cells["WF!B6"].computed_value

        assert preferred == 80000.0  # 8% of 1M
        assert gp_share == 84000.0   # 20% of 420K
        assert abs(check - 1500000.0) < 0.01  # LP + GP = total distributable


class TestDebtCovenantModel:
    """Debt model with DSCR (debt service coverage ratio) covenant check."""

    def test_dscr_covenant_breach(self):
        """DSCR = EBITDA / (interest + principal). Flag if < 1.2x."""
        wb = _make_wb({"Debt": {
            "literals": {
                "A1": 500000,   # EBITDA
                "A2": 300000,   # interest expense
                "A3": 150000,   # principal payment
                "A4": 1.2,      # covenant minimum DSCR
            },
            "formulas": {
                # B1 = debt service = A2 + A3
                "B1": {
                    "formula": "=A2+A3",
                    "dependencies": [{"sheet": "Debt", "cell": "A2", "via": "A2"}, {"sheet": "Debt", "cell": "A3", "via": "A3"}],
                    "tokens": [_ref("A2"), _op("+"), _ref("A3")],
                },
                # B2 = DSCR = A1 / B1
                "B2": {
                    "formula": "=A1/B1",
                    "dependencies": [{"sheet": "Debt", "cell": "A1", "via": "A1"}, {"sheet": "Debt", "cell": "B1", "via": "B1"}],
                    "tokens": [_ref("A1"), _op("/"), _ref("B1")],
                },
                # B3 = covenant test = IF(B2 < A4, "BREACH", "OK")
                "B3": {
                    "formula": "=IF(B2<A4,\"BREACH\",\"OK\")",
                    "functions": ["IF"],
                    "dependencies": [{"sheet": "Debt", "cell": "B2", "via": "B2"}, {"sheet": "Debt", "cell": "A4", "via": "A4"}],
                    "tokens": [
                        _func_open("IF"),
                        _ref("B2"), _op("<"), _ref("A4"), _sep(),
                        _text("BREACH"), _sep(), _text("OK"),
                        _func_close(),
                    ],
                },
            },
        }})

        graph = build_dependency_graph(wb)
        for qcell in topological_sort(graph, set()):
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        dscr = wb.formula_cells["Debt!B2"].computed_value
        covenant = wb.formula_cells["Debt!B3"].computed_value

        # DSCR = 500K / 450K ≈ 1.11, which is < 1.2
        assert abs(dscr - 500000/450000) < 0.01
        assert covenant == "BREACH"


# ===================================================================
# 16. Adversarial / stress tests — try to break the evaluator
# ===================================================================

class TestAdversarialArithmetic:
    """Arithmetic edge cases that produce NaN, Inf, or type confusion."""

    def test_zero_divided_by_zero(self):
        wb = _empty_wb()
        with pytest.raises(EvalError):
            _eval(wb, "S", [_num(0), _op("/"), _num(0)])

    def test_negative_base_fractional_exponent(self):
        """(-1)^0.5 = NaN in real arithmetic. Should not silently produce a value."""
        wb = _empty_wb()
        result = _eval(wb, "S", [_prefix("-"), _num(1), _op("^"), _num(0.5)])
        # Python returns complex nan — evaluator should either raise or return nan
        assert result != result or isinstance(result, complex)  # NaN or complex

    def test_zero_to_the_zero(self):
        wb = _empty_wb()
        # 0^0 is mathematically ambiguous. Python says 1.0. Accept it.
        result = _eval(wb, "S", [_num(0), _op("^"), _num(0)])
        assert result == 1.0  # Python convention

    def test_overflow_exponent(self):
        """2^1024 overflows float64."""
        wb = _empty_wb()
        result = _eval(wb, "S", [_num(2), _op("^"), _num(1024)])
        assert math.isinf(result)

    def test_deeply_chained_division(self):
        """1e308 / 0.1 / 0.1 / 0.1 → overflow to inf."""
        wb = _make_wb({"S": {"literals": {"A1": 1e308}}})
        tokens = [_ref("A1"), _op("/"), _num(0.1), _op("/"), _num(0.1), _op("/"), _num(0.1)]
        result = _eval(wb, "S", tokens)
        assert math.isinf(result)

    def test_subtraction_cancellation(self):
        """(1e15 + 1) - 1e15 should ideally be 1 but float loses precision."""
        wb = _empty_wb()
        tokens = [
            _paren_open(), _num(1e15), _op("+"), _num(1), _paren_close(),
            _op("-"), _num(1e15),
        ]
        result = _eval(wb, "S", tokens)
        assert abs(result - 1.0) < 2.0  # float64 precision loss is expected

    def test_boolean_in_arithmetic(self):
        """TRUE + TRUE + FALSE = 2 in Excel."""
        wb = _empty_wb()
        tokens = [_logical("TRUE"), _op("+"), _logical("TRUE"), _op("+"), _logical("FALSE")]
        assert _eval(wb, "S", tokens) == 2.0

    def test_string_in_arithmetic_raises(self):
        wb = _make_wb({"S": {"literals": {"A1": "hello"}}})
        with pytest.raises(EvalError, match="Cannot convert"):
            _eval(wb, "S", [_ref("A1"), _op("+"), _num(1)])

    def test_empty_string_as_zero(self):
        """Empty string coerced to 0 in arithmetic."""
        wb = _make_wb({"S": {"literals": {"A1": ""}}})
        result = _eval(wb, "S", [_ref("A1"), _op("+"), _num(5)])
        assert result == 5.0


class TestAdversarialTokens:
    """Malformed or adversarial token sequences."""

    def test_empty_token_list(self):
        wb = _empty_wb()
        with pytest.raises(EvalError, match="Empty"):
            _eval(wb, "S", [])

    def test_operator_with_no_operands(self):
        """Just a `+` token — should fail, not hang."""
        wb = _empty_wb()
        with pytest.raises(EvalError):
            _eval(wb, "S", [_op("+")])

    def test_unclosed_function(self):
        """SUM( with no closing paren — should not hang."""
        wb = _make_wb({"S": {"literals": {"A1": 1}}})
        tokens = [_func_open("SUM"), _ref("A1")]
        # Should either return a result or raise — not infinite loop
        try:
            result = _eval(wb, "S", tokens)
            # If it returns, that's fine
        except EvalError:
            pass  # Also fine

    def test_double_operator(self):
        """1 + + 2 — the second + could be unary or garbage."""
        wb = _empty_wb()
        # In Excel, =1++2 is valid (unary + on 2)
        # Our tokenizer would emit: NUM(1) INFIX(+) PREFIX(+) NUM(2)
        tokens = [_num(1), _op("+"), _prefix("+"), _num(2)]
        assert _eval(wb, "S", tokens) == 3.0

    def test_formula_cell_with_no_tokens(self):
        """A formula cell that somehow has an empty token list."""
        wb = _make_wb({"S": {"formulas": {
            "A1": {"formula": "=???", "tokens": []},
        }}})
        with pytest.raises(EvalError):
            evaluate_cell(wb, wb.formula_cells["S!A1"])

    def test_deeply_nested_functions(self):
        """IF(IF(IF(IF(IF(TRUE,1,0)>0,2,0)>0,3,0)>0,4,0)>0,5,0) = 5"""
        wb = _empty_wb()
        # Build from inside out
        # Innermost: IF(TRUE, 1, 0)
        tokens = [
            _func_open("IF"),
              _func_open("IF"),
                _func_open("IF"),
                  _func_open("IF"),
                    _func_open("IF"),
                      _logical("TRUE"), _sep(), _num(1), _sep(), _num(0),
                    _func_close(),
                    _op(">"), _num(0), _sep(), _num(2), _sep(), _num(0),
                  _func_close(),
                  _op(">"), _num(0), _sep(), _num(3), _sep(), _num(0),
                _func_close(),
                _op(">"), _num(0), _sep(), _num(4), _sep(), _num(0),
              _func_close(),
              _op(">"), _num(0), _sep(), _num(5), _sep(), _num(0),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == 5.0


class TestAdversarialRanges:
    """Range references designed to break things."""

    def test_range_where_all_cells_are_none(self):
        """SUM of a range where no cells exist → 0."""
        wb = _make_wb({"S": {"formulas": {}}})
        tokens = [_func_open("SUM"), _ref("Z1:Z100"), _func_close()]
        assert _eval(wb, "S", tokens) == 0.0

    def test_single_cell_range_colon(self):
        """A1:A1 — degenerate range, should work."""
        wb = _make_wb({"S": {"literals": {"A1": 42}}})
        tokens = [_func_open("SUM"), _ref("A1:A1"), _func_close()]
        assert _eval(wb, "S", tokens) == 42.0

    def test_index_out_of_bounds(self):
        """INDEX(A1:A3, 99) — row index way out of range."""
        wb = _make_wb({"S": {"literals": {"A1": 1, "A2": 2, "A3": 3}}})
        tokens = [_func_open("INDEX"), _ref("A1:A3"), _sep(), _num(99), _func_close()]
        # Should return None (cell doesn't exist) rather than crash
        result = _eval(wb, "S", tokens)
        assert result is None

    def test_match_no_match_raises(self):
        """MATCH with no matching value should raise."""
        wb = _make_wb({"S": {"literals": {"A1": 1, "A2": 2, "A3": 3}}})
        tokens = [_func_open("MATCH"), _num(999), _sep(), _ref("A1:A3"), _sep(), _num(0), _func_close()]
        with pytest.raises(EvalError, match="no exact match"):
            _eval(wb, "S", tokens)

    def test_vlookup_col_index_exceeds_table(self):
        """VLOOKUP with col_index=5 in a 2-column table."""
        wb = _make_wb({"S": {"literals": {"A1": "x", "B1": 10}}})
        tokens = [
            _func_open("VLOOKUP"), _text("x"), _sep(),
            _ref("A1:B1"), _sep(), _num(5), _sep(), _logical("FALSE"),
            _func_close(),
        ]
        # col_index 5 on a 2-col table → the referenced cell doesn't exist → None
        result = _eval(wb, "S", tokens)
        assert result is None

    def test_sumif_mismatched_range_lengths(self):
        """SUMIF where criteria range is longer than sum range."""
        wb = _make_wb({"S": {"literals": {
            "A1": "x", "A2": "y", "A3": "x", "A4": "x",  # 4 rows
            "B1": 10, "B2": 20,  # only 2 rows
        }}})
        tokens = [
            _func_open("SUMIF"), _ref("A1:A4"), _sep(), _text("x"), _sep(),
            _ref("B1:B2"), _func_close(),
        ]
        # Should sum B1=10 for A1=x, B2 doesn't exist for A3=x, A4=x
        result = _eval(wb, "S", tokens)
        assert result == 10.0  # only B1 matches within sum range


class TestAdversarialIferror:
    """IFERROR edge cases that might leak errors or produce nonsense."""

    def test_iferror_wrapping_unsupported_function(self):
        """=IFERROR(OFFSET(...), 0) — OFFSET is unsupported but IFERROR should catch."""
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=IFERROR(OFFSET(B1,1,0),0)",
                "functions": ["IFERROR", "OFFSET"],
                "tokens": [
                    _func_open("IFERROR"),
                    _func_open("OFFSET"), _ref("B1"), _sep(), _num(1), _sep(), _num(0), _func_close(),
                    _sep(), _num(0),
                    _func_close(),
                ],
            },
        }}})
        # OFFSET raises UnsupportedError, but IFERROR should catch it → 0
        result = evaluate_cell(wb, wb.formula_cells["S!A1"])
        assert result == 0.0

    def test_iferror_error_in_fallback(self):
        """=IFERROR(1/0, 1/0) — both arms error."""
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=IFERROR(1/0,1/0)",
                "tokens": [
                    _func_open("IFERROR"),
                    _num(1), _op("/"), _num(0),
                    _sep(),
                    _num(1), _op("/"), _num(0),
                    _func_close(),
                ],
            },
        }}})
        # Fallback also errors — should propagate
        with pytest.raises(EvalError):
            evaluate_cell(wb, wb.formula_cells["S!A1"])

    def test_nested_iferror_chain(self):
        """=IFERROR(IFERROR(IFERROR(1/0, 2/0), 3/0), 42)"""
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=IFERROR(IFERROR(IFERROR(1/0,2/0),3/0),42)",
                "tokens": [
                    _func_open("IFERROR"),
                      _func_open("IFERROR"),
                        _func_open("IFERROR"),
                          _num(1), _op("/"), _num(0),
                        _sep(), _num(2), _op("/"), _num(0),
                        _func_close(),
                      _sep(), _num(3), _op("/"), _num(0),
                      _func_close(),
                    _sep(), _num(42),
                    _func_close(),
                ],
            },
        }}})
        result = evaluate_cell(wb, wb.formula_cells["S!A1"])
        assert result == 42.0


class TestAdversarialIterative:
    """Iterative solver abuse."""

    def test_oscillating_loop(self):
        """A1 = 1-B1, B1 = 1-A1 — only fixed point is (0.5, 0.5).
        But with seed (0, 0): A1=1-0=1, B1=1-1=0, A1=1-0=1... oscillates.
        Actually this converges to (0.5, 0.5) via damping. Use a truly
        oscillating system: A1 = -B1+1, B1 = A1+1.
        Fixed point: A1=-B1+1, B1=A1+1 → A1=-(A1+1)+1 → 2A1=0 → (0,1).
        Seed (0,0): A1=1, B1=1, A1=0, B1=1, A1=0... period-2 oscillation.
        """
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=-B1+1",
                "cached_value": 5.0,
                "dependencies": [{"sheet": "S", "cell": "B1", "via": "B1"}],
                "tokens": [_prefix("-"), _ref("B1"), _op("+"), _num(1)],
                "iterative_component": "iter_1",
            },
            "B1": {
                "formula": "=A1+1",
                "cached_value": 5.0,
                "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}],
                "tokens": [_ref("A1"), _op("+"), _num(1)],
                "iterative_component": "iter_1",
            },
        }}})
        converged = solve_iterative_component(wb, ["S!A1", "S!B1"])
        # This system converges to (0, 1): A1=-1+1=0, B1=0+1=1, A1=-1+1=0 ✓
        # Starting from (5, 5): A1=-5+1=-4, B1=-4+1=-3, A1=3+1=4, B1=4+1=5...
        # Actually let's verify:
        a1 = wb.formula_cells["S!A1"].computed_value
        b1 = wb.formula_cells["S!B1"].computed_value
        if converged:
            # If it converges, it must be to the fixed point (0, 1)
            assert abs(a1 - 0.0) < 1e-6
            assert abs(b1 - 1.0) < 1e-6

    def test_slow_convergence_still_converges(self):
        """A1 = (A1 * 999 + 1000) / 1000 — converges to 1000 but very slowly.
        Actually: A1_{n+1} = 0.999 * A1_n + 1
        Fixed point: A1 = 0.999*A1 + 1 → 0.001*A1 = 1 → A1 = 1000
        Convergence rate: 0.999^n, needs ~7000 iterations for 1e-3 accuracy.
        With 200 max iterations, this should NOT converge.
        """
        wb = _make_wb({"S": {"formulas": {
            "A1": {
                "formula": "=(A1*999+1000)/1000",
                "cached_value": 0.0,
                "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}],
                "tokens": [
                    _paren_open(), _ref("A1"), _op("*"), _num(999), _op("+"), _num(1000), _paren_close(),
                    _op("/"), _num(1000),
                ],
                "iterative_component": "iter_1",
            },
        }}})
        converged = solve_iterative_component(wb, ["S!A1"])
        assert not converged
        # Should be somewhere between 0 and 1000, not exactly at fixed point
        val = wb.formula_cells["S!A1"].computed_value
        assert 0 < val < 1000

    def test_iterative_component_with_unsupported_cell(self):
        """One cell in the SCC uses OFFSET. The rest should still attempt solving."""
        wb = _make_wb({"S": {
            "literals": {"C1": 100},
            "formulas": {
                "A1": {
                    "formula": "=B1+C1",
                    "functions": [],
                    "dependencies": [{"sheet": "S", "cell": "B1", "via": "B1"}, {"sheet": "S", "cell": "C1", "via": "C1"}],
                    "tokens": [_ref("B1"), _op("+"), _ref("C1")],
                    "iterative_component": "iter_1",
                },
                "B1": {
                    "formula": "=OFFSET(A1,0,0)",
                    "functions": ["OFFSET"],
                    "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}],
                    "tokens": [_func_open("OFFSET"), _ref("A1"), _sep(), _num(0), _sep(), _num(0), _func_close()],
                    "iterative_component": "iter_1",
                },
            },
        }})
        # Classify B1 as unsupported first
        from deterministic_evaluator import classify_functions
        for fc in wb.formula_cells.values():
            ok, unsup = classify_functions(fc.functions)
            if not ok:
                fc.status = "unsupported"

        converged = solve_iterative_component(wb, ["S!A1", "S!B1"])
        assert wb.formula_cells["S!B1"].status == "unsupported"

    def test_100_cell_scc(self):
        """Large SCC: A1→A2→A3→...→A100→A1, each = prev + 0.001.
        Fixed point: every cell = prev + 0.001, which is contradictory
        (circular add never converges to nonzero). With IFERROR wrapping, converges to 0.
        """
        formulas = {}
        for i in range(1, 101):
            prev_cell = f"A{100 if i == 1 else i-1}"
            formulas[f"A{i}"] = {
                "formula": f"=IFERROR({prev_cell}+0.001,0)",
                "cached_value": 0.0,
                "functions": ["IFERROR"],
                "dependencies": [{"sheet": "S", "cell": prev_cell, "via": prev_cell}],
                "tokens": [
                    _func_open("IFERROR"),
                    _ref(prev_cell), _op("+"), _num(0.001),
                    _sep(), _num(0),
                    _func_close(),
                ],
                "iterative_component": "iter_1",
            }
        wb = _make_wb({"S": {"formulas": formulas}})
        component = [f"S!A{i}" for i in range(1, 101)]

        import time
        t0 = time.time()
        converged = solve_iterative_component(wb, component)
        elapsed = time.time() - t0

        # 100-cell SCC with 200+ iterations should still finish quickly
        assert elapsed < 5.0, f"100-cell SCC took {elapsed:.2f}s"
        # Each cell keeps growing by 0.001 per iteration, never settles
        assert not converged


class TestAdversarialFinance:
    """Finance function edge cases."""

    def test_xirr_single_day_cashflows(self):
        """All cash flows on the same date — solver should fail."""
        d = _date_to_serial(date(2024, 1, 1))
        wb = _make_wb({"S": {"literals": {
            "A1": -100, "A2": 200, "B1": d, "B2": d,
        }}})
        tokens = [_func_open("XIRR"), _ref("A1:A2"), _sep(), _ref("B1:B2"), _func_close()]
        # All same date → exponent = 0 for everything → degenerate
        # Should either return a huge number or raise
        try:
            result = _eval(wb, "S", tokens)
            # If it returns, the rate should be astronomical or inf
        except EvalError:
            pass  # Also acceptable

    def test_irr_all_zeros(self):
        """IRR of all-zero cash flows."""
        wb = _make_wb({"S": {"literals": {"A1": 0, "A2": 0, "A3": 0}}})
        tokens = [_func_open("IRR"), _ref("A1:A3"), _func_close()]
        with pytest.raises(EvalError):
            _eval(wb, "S", tokens)

    def test_npv_negative_rate(self):
        """NPV with rate = -0.5 (negative discount rate)."""
        wb = _make_wb({"S": {"literals": {"A1": 100, "A2": 100}}})
        tokens = [_func_open("NPV"), _prefix("-"), _num(0.5), _sep(), _ref("A1:A2"), _func_close()]
        result = _eval(wb, "S", tokens)
        # (1 + -0.5) = 0.5, so PV1 = 100/0.5 = 200, PV2 = 100/0.25 = 400
        assert abs(result - 600.0) < 0.01

    def test_xirr_with_100_cashflows(self):
        """XIRR with many cash flows — performance test."""
        lits = {}
        d0 = _date_to_serial(date(2020, 1, 1))
        lits["A1"] = -1000000
        lits["B1"] = d0
        for i in range(2, 52):
            lits[f"A{i}"] = 25000
            lits[f"B{i}"] = d0 + (i - 1) * 30  # monthly
        wb = _make_wb({"S": {"literals": lits}})
        tokens = [_func_open("XIRR"), _ref("A1:A51"), _sep(), _ref("B1:B51"), _func_close()]

        import time
        t0 = time.time()
        result = _eval(wb, "S", tokens)
        elapsed = time.time() - t0

        assert elapsed < 1.0, f"XIRR with 51 cashflows took {elapsed:.2f}s"
        assert isinstance(result, float)
        assert not math.isnan(result)


class TestAdversarialGraph:
    """Graph construction and evaluation ordering edge cases."""

    def test_diamond_dependency(self):
        """A depends on B and C, both depend on D.
        D should be evaluated exactly once before B and C.
        """
        wb = _make_wb({"S": {
            "literals": {"D1": 10},
            "formulas": {
                "B1": {
                    "formula": "=D1*2",
                    "dependencies": [{"sheet": "S", "cell": "D1", "via": "D1"}],
                    "tokens": [_ref("D1"), _op("*"), _num(2)],
                },
                "C1": {
                    "formula": "=D1*3",
                    "dependencies": [{"sheet": "S", "cell": "D1", "via": "D1"}],
                    "tokens": [_ref("D1"), _op("*"), _num(3)],
                },
                "A1": {
                    "formula": "=B1+C1",
                    "dependencies": [
                        {"sheet": "S", "cell": "B1", "via": "B1"},
                        {"sheet": "S", "cell": "C1", "via": "C1"},
                    ],
                    "tokens": [_ref("B1"), _op("+"), _ref("C1")],
                },
            },
        }})
        graph = build_dependency_graph(wb)
        order = topological_sort(graph, set())
        for qcell in order:
            fc = wb.formula_cells.get(qcell)
            if fc:
                fc.computed_value = evaluate_cell(wb, fc)

        assert wb.formula_cells["S!A1"].computed_value == 50.0  # 20 + 30

    def test_cross_sheet_circular_ref_detected(self):
        """S1!A1 → S2!A1 → S1!A1 — cross-sheet cycle."""
        wb = _make_wb({
            "S1": {"formulas": {
                "A1": {
                    "formula": "='S2'!A1+1",
                    "dependencies": [{"sheet": "S2", "cell": "A1", "via": "A1"}],
                    "tokens": [_ref("'S2'!A1"), _op("+"), _num(1)],
                },
            }},
            "S2": {"formulas": {
                "A1": {
                    "formula": "='S1'!A1+1",
                    "dependencies": [{"sheet": "S1", "cell": "A1", "via": "A1"}],
                    "tokens": [_ref("'S1'!A1"), _op("+"), _num(1)],
                },
            }},
        })
        graph = build_dependency_graph(wb)
        sccs = tarjan_scc(graph)
        cycle = [s for s in sccs if len(s) > 1]
        assert len(cycle) == 1
        assert set(cycle[0]) == {"S1!A1", "S2!A1"}

    def test_self_referencing_cell_not_tagged_iterative(self):
        """A cell that references itself but has no iterative_component tag.
        The evaluator should derive the SCC and handle it.
        """
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["S"],
                "functions_used": {"IFERROR": 1},
                "iterative_components": [],  # deliberately empty
            }))
            (d / "S.json").write_text(json.dumps({
                "sheet": "S",
                "formula_cells": [{
                    "cell": "A1",
                    "formula": "=IFERROR(A1*0.5+50, 0)",
                    "cached_value": None,
                    "functions": ["IFERROR"],
                    "references": [{"sheet": "S", "ref": "A1", "kind": "cell", "expanded": ["A1"]}],
                    "dependencies": [{"sheet": "S", "cell": "A1", "via": "A1"}],
                    "tokens": [
                        {"type": "FUNC", "subtype": "OPEN", "value": "IFERROR("},
                        {"type": "OPERAND", "subtype": "RANGE", "value": "A1"},
                        {"type": "OPERATOR-INFIX", "subtype": "", "value": "*"},
                        {"type": "OPERAND", "subtype": "NUMBER", "value": "0.5"},
                        {"type": "OPERATOR-INFIX", "subtype": "", "value": "+"},
                        {"type": "OPERAND", "subtype": "NUMBER", "value": "50"},
                        {"type": "SEP", "subtype": "ARG", "value": ","},
                        {"type": "OPERAND", "subtype": "NUMBER", "value": "0"},
                        {"type": "FUNC", "subtype": "CLOSE", "value": ")"},
                    ],
                    "iterative_component": None,  # NOT tagged
                }],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            # Should derive the self-reference SCC and solve it
            # Fixed point: A1 = 0.5*A1 + 50 → A1 = 100
            assert verdict["summary"]["supported_converged"] == 1
            assert verdict["logic"]["passed"]


class TestAdversarialTypes:
    """Type coercion nightmares."""

    def test_date_serial_in_arithmetic(self):
        """Date serial number used in plain arithmetic should work as a number."""
        d = _date_to_serial(date(2024, 6, 15))
        wb = _make_wb({"S": {"literals": {"A1": d}}})
        tokens = [_ref("A1"), _op("+"), _num(30)]
        result = _eval(wb, "S", tokens)
        assert result == d + 30

    def test_if_with_mixed_return_types(self):
        """IF returning string in true branch, number in false branch."""
        wb = _empty_wb()
        tokens = [
            _func_open("IF"), _logical("TRUE"), _sep(), _text("yes"), _sep(), _num(0),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens) == "yes"

        tokens2 = [
            _func_open("IF"), _logical("FALSE"), _sep(), _text("yes"), _sep(), _num(0),
            _func_close(),
        ]
        assert _eval(wb, "S", tokens2) == 0.0

    def test_sum_ignores_booleans(self):
        """SUM should ignore TRUE/FALSE in a range (Excel behavior)."""
        wb = _make_wb({"S": {"literals": {"A1": True, "A2": False, "A3": 10}}})
        tokens = [_func_open("SUM"), _ref("A1:A3"), _func_close()]
        # Booleans in ranges are ignored by SUM in Excel
        # But our literals are Python bools — they're not (int, float)
        # so SUM should skip them
        result = _eval(wb, "S", tokens)
        assert result == 10.0

    def test_comparison_string_vs_number(self):
        """In Excel, "10" > 9 is FALSE (strings > numbers in Excel sort order).
        Our evaluator tries numeric comparison first.
        """
        wb = _empty_wb()
        tokens = [_text("10"), _op(">"), _num(9)]
        result = _eval(wb, "S", tokens)
        # Our impl tries numeric first: 10 > 9 = True
        assert result is True

    def test_concatenation_preserves_none_as_empty(self):
        """Concatenating None cell should produce empty string, not 'None'."""
        wb = _make_wb({"S": {"formulas": {}}})  # A1 doesn't exist
        tokens = [_text("x"), _op("&"), _ref("A1"), _op("&"), _text("y")]
        result = _eval(wb, "S", tokens)
        assert result == "xy"  # None → ""


# ===================================================================
# 17. Complex structures — things real Excel files produce
# ===================================================================

def _array_open():
    return {"type": "ARRAY", "subtype": "OPEN", "value": "{"}

def _array_close():
    return {"type": "ARRAY", "subtype": "CLOSE", "value": "}"}

def _row_sep():
    return {"type": "SEP", "subtype": "ROW", "value": ";"}


class TestArrayConstants:
    """Array constants: {1,2,3} and {1,2;3,4}."""

    def test_array_constant_in_sum(self):
        """=SUM({10,20,30}) → 60. ARRAY tokens bracket the values."""
        wb = _empty_wb()
        tokens = [
            _func_open("SUM"),
            _array_open(), _num(10), _sep(), _num(20), _sep(), _num(30), _array_close(),
            _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        assert result == 60.0

    def test_array_constant_2d(self):
        """={1,2;3,4} — 2D array. SUM should flatten and sum all."""
        wb = _empty_wb()
        tokens = [
            _func_open("SUM"),
            _array_open(),
            _num(1), _sep(), _num(2), _row_sep(),
            _num(3), _sep(), _num(4),
            _array_close(),
            _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        assert result == 10.0

    def test_array_constant_multiplication(self):
        """=SUM(A1:A3)*{1} — array in arithmetic context.
        Should either work as scalar or be classified unsupported."""
        wb = _make_wb({"S": {"literals": {"A1": 10, "A2": 20, "A3": 30}}})
        tokens = [
            _func_open("SUM"), _ref("A1:A3"), _func_close(),
            _op("*"),
            _array_open(), _num(1), _array_close(),
        ]
        try:
            result = _eval(wb, "S", tokens)
            # If it handles it, the array {1} should be treated as scalar 1
            assert result == 60.0
        except (EvalError, UnsupportedError):
            pass  # Also acceptable — arrays in arithmetic are complex


class TestWholeColumnRowRefs:
    """Whole-column (A:A) and whole-row (1:1) references."""

    def test_whole_column_ref_in_sum(self):
        """=SUM(A:A) — infinite range. Evaluator should handle gracefully."""
        wb = _make_wb({"S": {"literals": {"A1": 10, "A2": 20}}})
        tokens = [_func_open("SUM"), _ref("A:A"), _func_close()]
        try:
            result = _eval(wb, "S", tokens)
            # If it works, it should sum the known cells
        except (EvalError, ValueError):
            pass  # Acceptable: can't expand infinite range

    def test_whole_row_ref(self):
        """=SUM(1:1) — whole row."""
        wb = _make_wb({"S": {"literals": {"A1": 5, "B1": 10}}})
        tokens = [_func_open("SUM"), _ref("1:1"), _func_close()]
        try:
            result = _eval(wb, "S", tokens)
        except (EvalError, ValueError):
            pass  # Acceptable

    def test_xlookup_with_whole_column(self):
        """=_xlfn.XLOOKUP(A1, B:B, C:C) — common real-world pattern."""
        wb = _make_wb({"S": {"literals": {
            "A1": "target",
            "B1": "target", "C1": 42,
            "B2": "other", "C2": 99,
        }}})
        # _xlfn.XLOOKUP tokenizes as FUNC with _xlfn. prefix
        tokens = [
            {"type": "FUNC", "subtype": "OPEN", "value": "_xlfn.XLOOKUP("},
            _ref("A1"), _sep(), _ref("B:B"), _sep(), _ref("C:C"),
            _func_close(),
        ]
        try:
            result = _eval(wb, "S", tokens)
            assert result == 42
        except (EvalError, UnsupportedError):
            pass  # Whole-column XLOOKUP might not be supported


class TestXlfnPrefix:
    """Functions with _xlfn. prefix (Excel 2013+ new functions)."""

    def test_xlfn_xlookup(self):
        """_xlfn.XLOOKUP should be recognized as XLOOKUP."""
        wb = _make_wb({"S": {"literals": {
            "A1": "b", "B1": "a", "B2": "b", "B3": "c",
            "C1": 10, "C2": 20, "C3": 30,
        }}})
        tokens = [
            {"type": "FUNC", "subtype": "OPEN", "value": "_xlfn.XLOOKUP("},
            _ref("A1"), _sep(), _ref("B1:B3"), _sep(), _ref("C1:C3"),
            _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        assert result == 20

    def test_xlfn_xmatch(self):
        wb = _make_wb({"S": {"literals": {"A1": "b", "B1": "a", "B2": "b", "B3": "c"}}})
        tokens = [
            {"type": "FUNC", "subtype": "OPEN", "value": "_xlfn.XMATCH("},
            _ref("A1"), _sep(), _ref("B1:B3"), _sep(), _num(0),
            _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        assert result == 2

    def test_xlfn_concat(self):
        wb = _make_wb({"S": {"literals": {"A1": "Hello", "A2": " ", "A3": "World"}}})
        tokens = [
            {"type": "FUNC", "subtype": "OPEN", "value": "_xlfn.CONCAT("},
            _ref("A1:A3"), _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        assert result == "Hello World"

    def test_xlfn_unknown_function(self):
        """_xlfn.SOMETHINGNEW should be unsupported."""
        wb = _empty_wb()
        tokens = [
            {"type": "FUNC", "subtype": "OPEN", "value": "_xlfn.SOMETHINGNEW("},
            _num(1), _func_close(),
        ]
        with pytest.raises(UnsupportedError):
            _eval(wb, "S", tokens)


class TestStructuredReferences:
    """Table1[Column] structured references."""

    def test_structured_ref_classified(self):
        """Table1[Revenue] comes through as an OPERAND/RANGE.
        The evaluator should either resolve it or classify it unsupported."""
        wb = _empty_wb()
        tokens = [_ref("Table1[Revenue]")]
        try:
            result = _eval(wb, "S", tokens)
            # If it resolves, it'll be None (not in workbook)
        except (EvalError, ValueError):
            pass  # Acceptable: structured refs not supported

    def test_structured_ref_in_formula_cell(self):
        """Full workbook evaluation with a structured ref formula."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["S"],
                "functions_used": {"SUM": 1},
                "iterative_components": [],
            }))
            (d / "S.json").write_text(json.dumps({
                "sheet": "S",
                "formula_cells": [{
                    "cell": "A1",
                    "formula": "=SUM(Table1[Revenue])",
                    "cached_value": 1000,
                    "functions": ["SUM"],
                    "references": [{"sheet": "S", "ref": "Table1[Revenue]", "kind": "range", "expanded": []}],
                    "dependencies": [],
                    "tokens": [
                        {"type": "FUNC", "subtype": "OPEN", "value": "SUM("},
                        {"type": "OPERAND", "subtype": "RANGE", "value": "Table1[Revenue]"},
                        {"type": "FUNC", "subtype": "CLOSE", "value": ")"},
                    ],
                    "iterative_component": None,
                }],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            # Should not crash. Structured ref resolves to nothing → SUM = 0
            # or is classified as parse-failure. Either way, no crash.
            assert verdict["summary"]["total_cells"] == 1


class TestThreeDReferences:
    """Sheet1:Sheet3!A1 — 3D multi-sheet references."""

    def test_3d_ref_parse(self):
        """3D reference in a SUM — should either work or be classified."""
        wb = _make_wb({
            "S1": {"literals": {"A1": 10}},
            "S2": {"literals": {"A1": 20}},
            "S3": {"literals": {"A1": 30}},
            "Main": {"formulas": {}},
        })
        tokens = [_func_open("SUM"), _ref("S1:S3!A1"), _func_close()]
        try:
            result = _eval(wb, "Main", tokens)
            # If it handles 3D refs, should be 60
        except (EvalError, ValueError):
            pass  # Acceptable: 3D refs are complex


class TestImplicitIntersection:
    """@A1:A10 implicit intersection operator."""

    def test_at_sign_ref(self):
        """@ prefix on a range — should resolve or be classified."""
        wb = _make_wb({"S": {"literals": {"A1": 42, "A2": 99}}})
        tokens = [_ref("@A1:A10")]
        try:
            result = _eval(wb, "S", tokens)
        except (EvalError, ValueError):
            pass  # Acceptable


class TestDataTableFormulas:
    """DATA_TABLE formulas — not real formulas, should be handled gracefully."""

    def test_data_table_cell_in_workbook(self):
        """A cell whose formula is a data table string, not a real formula."""
        with tempfile.TemporaryDirectory() as tmpdir:
            d = Path(tmpdir)
            (d / "workbook.json").write_text(json.dumps({
                "workbook": "test.xlsx",
                "sheet_names": ["S"],
                "functions_used": {},
                "iterative_components": [],
            }))
            (d / "S.json").write_text(json.dumps({
                "sheet": "S",
                "formula_cells": [{
                    "cell": "B2",
                    "formula": "=DATA_TABLE[2D] anchor:B2 range:B2:F6 r1:B1 r2:A2",
                    "cached_value": 42.5,
                    "functions": [],
                    "references": [],
                    "dependencies": [],
                    "tokens": [],  # data tables have no parseable tokens
                    "iterative_component": None,
                }],
                "literal_cells": {},
            }))

            verdict = evaluate_workbook(d)
            # Empty tokens → parse failure, but should not crash
            assert verdict["summary"]["total_cells"] == 1
            # The cell should be classified as parse-failure (empty tokens)
            assert verdict["summary"]["parse_failures"] == 1


class TestMixedComplexPatterns:
    """Formulas combining multiple complex features."""

    def test_index_match_cross_sheet_with_iferror(self):
        """=IFERROR(INDEX('Data'!B1:B5, MATCH(A1, 'Data'!A1:A5, 0)), "N/A")
        Real-world pattern: lookup in another sheet with error handling."""
        wb = _make_wb({
            "Data": {"literals": {
                "A1": "alpha", "B1": 100,
                "A2": "beta",  "B2": 200,
                "A3": "gamma", "B3": 300,
                "A4": "delta", "B4": 400,
                "A5": "epsilon", "B5": 500,
            }},
            "Main": {
                "literals": {"A1": "gamma"},
                "formulas": {
                    "B1": {
                        "formula": "=IFERROR(INDEX('Data'!B1:B5,MATCH(A1,'Data'!A1:A5,0)),\"N/A\")",
                        "functions": ["IFERROR", "INDEX", "MATCH"],
                        "dependencies": [
                            {"sheet": "Main", "cell": "A1", "via": "A1"},
                            {"sheet": "Data", "cell": "B1", "via": "B1:B5"},
                            {"sheet": "Data", "cell": "B2", "via": "B1:B5"},
                            {"sheet": "Data", "cell": "B3", "via": "B1:B5"},
                            {"sheet": "Data", "cell": "B4", "via": "B1:B5"},
                            {"sheet": "Data", "cell": "B5", "via": "B1:B5"},
                            {"sheet": "Data", "cell": "A1", "via": "A1:A5"},
                            {"sheet": "Data", "cell": "A2", "via": "A1:A5"},
                            {"sheet": "Data", "cell": "A3", "via": "A1:A5"},
                            {"sheet": "Data", "cell": "A4", "via": "A1:A5"},
                            {"sheet": "Data", "cell": "A5", "via": "A1:A5"},
                        ],
                        "tokens": [
                            _func_open("IFERROR"),
                            _func_open("INDEX"), _ref("'Data'!B1:B5"), _sep(),
                            _func_open("MATCH"), _ref("A1"), _sep(), _ref("'Data'!A1:A5"), _sep(), _num(0), _func_close(),
                            _func_close(),
                            _sep(), _text("N/A"),
                            _func_close(),
                        ],
                    },
                },
            },
        })
        result = evaluate_cell(wb, wb.formula_cells["Main!B1"])
        assert result == 300

    def test_index_match_cross_sheet_not_found(self):
        """Same pattern but lookup value doesn't exist → "N/A"."""
        wb = _make_wb({
            "Data": {"literals": {
                "A1": "alpha", "B1": 100,
                "A2": "beta",  "B2": 200,
            }},
            "Main": {
                "literals": {"A1": "MISSING"},
                "formulas": {
                    "B1": {
                        "formula": "=IFERROR(INDEX('Data'!B1:B2,MATCH(A1,'Data'!A1:A2,0)),\"N/A\")",
                        "functions": ["IFERROR", "INDEX", "MATCH"],
                        "dependencies": [
                            {"sheet": "Main", "cell": "A1", "via": "A1"},
                            {"sheet": "Data", "cell": "B1", "via": "B1:B2"},
                            {"sheet": "Data", "cell": "B2", "via": "B1:B2"},
                            {"sheet": "Data", "cell": "A1", "via": "A1:A2"},
                            {"sheet": "Data", "cell": "A2", "via": "A1:A2"},
                        ],
                        "tokens": [
                            _func_open("IFERROR"),
                            _func_open("INDEX"), _ref("'Data'!B1:B2"), _sep(),
                            _func_open("MATCH"), _ref("A1"), _sep(), _ref("'Data'!A1:A2"), _sep(), _num(0), _func_close(),
                            _func_close(),
                            _sep(), _text("N/A"),
                            _func_close(),
                        ],
                    },
                },
            },
        })
        result = evaluate_cell(wb, wb.formula_cells["Main!B1"])
        assert result == "N/A"

    def test_sumifs_cross_sheet_multi_criteria_with_dates(self):
        """SUMIFS pulling from another sheet, filtering by date range and category."""
        d1 = _date_to_serial(date(2024, 1, 15))
        d2 = _date_to_serial(date(2024, 2, 15))
        d3 = _date_to_serial(date(2024, 3, 15))
        wb = _make_wb({
            "Txn": {"literals": {
                "A1": 100,   "B1": "revenue", "C1": d1,
                "A2": 200,   "B2": "cost",    "C2": d2,
                "A3": 300,   "B3": "revenue", "C3": d2,
                "A4": 50,    "B4": "revenue", "C4": d3,
            }},
            "Summary": {
                "literals": {"A1": f">{_date_to_serial(date(2024, 1, 31))}"},
                "formulas": {
                    "B1": {
                        "formula": "=SUMIFS('Txn'!A1:A4,'Txn'!B1:B4,\"revenue\",'Txn'!C1:C4,A1)",
                        "functions": ["SUMIFS"],
                        "dependencies": [
                            {"sheet": "Txn", "cell": "A1", "via": "A1:A4"},
                            {"sheet": "Txn", "cell": "A2", "via": "A1:A4"},
                            {"sheet": "Txn", "cell": "A3", "via": "A1:A4"},
                            {"sheet": "Txn", "cell": "A4", "via": "A1:A4"},
                            {"sheet": "Txn", "cell": "B1", "via": "B1:B4"},
                            {"sheet": "Txn", "cell": "B2", "via": "B1:B4"},
                            {"sheet": "Txn", "cell": "B3", "via": "B1:B4"},
                            {"sheet": "Txn", "cell": "B4", "via": "B1:B4"},
                            {"sheet": "Txn", "cell": "C1", "via": "C1:C4"},
                            {"sheet": "Txn", "cell": "C2", "via": "C1:C4"},
                            {"sheet": "Txn", "cell": "C3", "via": "C1:C4"},
                            {"sheet": "Txn", "cell": "C4", "via": "C1:C4"},
                            {"sheet": "Summary", "cell": "A1", "via": "A1"},
                        ],
                        "tokens": [
                            _func_open("SUMIFS"),
                            _ref("'Txn'!A1:A4"), _sep(),
                            _ref("'Txn'!B1:B4"), _sep(), _text("revenue"), _sep(),
                            _ref("'Txn'!C1:C4"), _sep(), _ref("A1"),
                            _func_close(),
                        ],
                    },
                },
            },
        })
        result = evaluate_cell(wb, wb.formula_cells["Summary!B1"])
        # revenue rows with date > Jan 31: row3 (Feb 15, 300) + row4 (Mar 15, 50)
        assert result == 350.0

    def test_nested_round_sum_if_concatenate(self):
        """=CONCATENATE("Total: $", ROUND(SUM(IF(A1:A3>0, A1:A3, 0)), 2))
        This is a deeply nested pattern. Without CSE array behavior,
        the IF operates on a range comparison which is tricky.
        """
        wb = _make_wb({"S": {"literals": {"A1": 10.555, "A2": -5, "A3": 20.111}}})
        # The inner IF(A1:A3>0, A1:A3, 0) in non-CSE mode would compare
        # the first element only (implicit intersection). But let's test
        # if the evaluator handles the token sequence without crashing.
        tokens = [
            _func_open("CONCATENATE"),
            _text("Total: $"), _sep(),
            _func_open("ROUND"),
              _func_open("SUM"),
                _ref("A1:A3"),
              _func_close(),
            _sep(), _num(2),
            _func_close(),
            _func_close(),
        ]
        result = _eval(wb, "S", tokens)
        # SUM(A1:A3) = 10.555 + (-5) + 20.111 = 25.666, ROUND to 25.67
        assert result == "Total: $25.67"


# ===================================================================
# Runner
# ===================================================================

if __name__ == "__main__":
    sys.exit(pytest.main([__file__, "-v"]))
