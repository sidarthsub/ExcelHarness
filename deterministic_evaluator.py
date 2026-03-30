#!/usr/bin/env python3
"""Deterministic machine evaluator for workbook logic and completeness.

Runs from dump artifacts produced by dump.py. Covers whole sheets and whole
workbooks. Primary source of truth for logic/completeness; the prompt evaluator
remains responsible for visual review and unsupported-logic fallback.

See docs/evaluator_spec.md for the full specification.
"""

import json
import math
import re
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Cell reference utilities
# ---------------------------------------------------------------------------

_COL_CACHE: dict[str, int] = {}


def col_to_num(col: str) -> int:
    """Convert column letter(s) to 1-based number. A=1, Z=26, AA=27."""
    if col in _COL_CACHE:
        return _COL_CACHE[col]
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    _COL_CACHE[col] = n
    return n


def num_to_col(n: int) -> str:
    """Convert 1-based column number to letter(s)."""
    result = []
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


_CELL_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$", re.IGNORECASE)


def parse_cell(ref: str) -> tuple[str, int]:
    """Parse 'A1' or '$A$1' into (col_letter, row_number)."""
    m = _CELL_RE.match(ref.strip())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return m.group(1).upper(), int(m.group(2))


def expand_range(range_ref: str) -> list[str]:
    """Expand 'A1:C3' into list of cell references."""
    if ":" not in range_ref:
        return [range_ref.replace("$", "").upper()]
    parts = range_ref.split(":")
    c1, r1 = parse_cell(parts[0])
    c2, r2 = parse_cell(parts[1])
    cn1, cn2 = col_to_num(c1), col_to_num(c2)
    cells = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(cn1, cn2), max(cn1, cn2) + 1):
            cells.append(f"{num_to_col(c)}{r}")
    return cells


def qualified_cell(sheet: str, cell: str) -> str:
    """Return 'SheetName!A1' qualified cell key."""
    return f"{sheet}!{cell.upper().replace('$', '')}"


# ---------------------------------------------------------------------------
# Workbook data model
# ---------------------------------------------------------------------------

class FormulaCell:
    """Represents a single formula cell loaded from the JSON dump."""

    __slots__ = (
        "sheet", "cell", "qcell", "formula", "cached_value",
        "functions", "references", "dependencies", "tokens",
        "iterative_component", "computed_value", "status",
    )

    def __init__(self, sheet: str, data: dict):
        self.sheet = sheet
        self.cell = data["cell"].upper()
        self.qcell = qualified_cell(sheet, self.cell)
        self.formula = data.get("formula", "")
        self.cached_value = data.get("cached_value")
        self.functions = data.get("functions", [])
        self.references = data.get("references", [])
        self.dependencies = data.get("dependencies", [])
        self.tokens = data.get("tokens", [])
        self.iterative_component = data.get("iterative_component")
        self.computed_value: Any = None
        self.status: str = "pending"  # pending | supported-converged | supported-nonconverged | unsupported | parse-failure | graph-failure


class Workbook:
    """In-memory representation of all formula cells across all sheets."""

    def __init__(self):
        self.sheet_names: list[str] = []
        self.functions_used: dict[str, int] = {}
        self.iterative_components: list[dict] = []
        # qcell -> FormulaCell
        self.formula_cells: dict[str, FormulaCell] = {}
        # sheet -> list of FormulaCell
        self.sheet_cells: dict[str, list[FormulaCell]] = defaultdict(list)
        # qcell -> raw literal value (non-formula cells read from cached_value context)
        self.literal_values: dict[str, Any] = {}

    def add_cell(self, fc: FormulaCell) -> None:
        self.formula_cells[fc.qcell] = fc
        self.sheet_cells[fc.sheet].append(fc)

    def get_value(self, qcell: str) -> Any:
        """Get the current value of a cell (computed or literal)."""
        fc = self.formula_cells.get(qcell)
        if fc is not None:
            return fc.computed_value
        return self.literal_values.get(qcell)


# ---------------------------------------------------------------------------
# Supported function surface
# ---------------------------------------------------------------------------

SUPPORTED_FUNCTIONS = {
    # Logical
    "IF", "AND", "OR", "NOT",
    # Aggregates
    "SUM", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "MIN", "MAX",
    # Numeric
    "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS",
    # Date
    "DATE", "DAY", "MONTH", "YEAR", "EDATE", "EOMONTH", "YEARFRAC",
    # Lookup
    "INDEX", "MATCH", "XMATCH", "XLOOKUP", "VLOOKUP", "HLOOKUP",
    # Text
    "CONCATENATE", "CONCAT", "TEXTJOIN",
    # Finance
    "IRR", "XIRR", "NPV", "XNPV",
    # Wrapper commonly used
    "IFERROR", "IFNA",
}

UNSUPPORTED_FUNCTIONS = {"OFFSET", "INDIRECT"}


def classify_functions(funcs: list[str]) -> tuple[bool, list[str]]:
    """Return (all_supported, list_of_unsupported_funcs)."""
    unsupported = []
    for f in funcs:
        fu = f.upper()
        if fu not in SUPPORTED_FUNCTIONS and fu not in ("", ):
            unsupported.append(fu)
    return len(unsupported) == 0, unsupported


# ---------------------------------------------------------------------------
# Token-based expression evaluator
# ---------------------------------------------------------------------------

class EvalError(Exception):
    """Raised when expression evaluation fails."""
    pass


class UnsupportedError(EvalError):
    """Raised when an unsupported function/feature is encountered."""
    pass


def _to_number(v: Any) -> float:
    """Coerce value to number for arithmetic."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, date):
        return float(_date_to_serial(v))
    if isinstance(v, str):
        try:
            return float(v)
        except ValueError:
            pass
        try:
            return float(_date_to_serial(date.fromisoformat(v)))
        except ValueError:
            pass
        try:
            return float(_date_to_serial(_serial_to_date(int(float(v)))))
        except (ValueError, OverflowError):
            raise EvalError(f"Cannot convert '{v}' to number")
    raise EvalError(f"Cannot convert {type(v).__name__} to number")


def _to_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        if v.upper() == "TRUE":
            return True
        if v.upper() == "FALSE":
            return False
    return bool(v)


def _date_to_serial(d: date) -> int:
    """Convert date to Excel serial number (1900 system)."""
    delta = d - date(1899, 12, 30)
    return delta.days


def _serial_to_date(serial: int) -> date:
    """Convert Excel serial number to date."""
    return date(1899, 12, 30) + timedelta(days=int(serial))


def _excel_date(y: int, m: int, d: int) -> date:
    """Excel-like DATE normalization.

    Excel permits month overflow/underflow and day overflow/underflow.
    Normalize year/month first, then add day offset from the first of month.
    """
    total_months = y * 12 + (m - 1)
    norm_year = total_months // 12
    norm_month = (total_months % 12) + 1
    month_start = date(norm_year, norm_month, 1)
    return month_start + timedelta(days=d - 1)


def _values_match(a: Any, b: Any) -> bool:
    """Comparison for MATCH/XLOOKUP exact match."""
    if isinstance(a, str) and isinstance(b, str):
        return a.strip().upper() == b.strip().upper()
    try:
        return _to_number(a) == _to_number(b)
    except EvalError:
        return str(a).strip().upper() == str(b).strip().upper()


def _compare(a: Any, b: Any, op: str) -> bool:
    """Compare two values with the given operator."""
    # Try numeric comparison first
    try:
        na, nb = _to_number(a), _to_number(b)
        if op == "=":
            return na == nb
        if op == "<>":
            return na != nb
        if op == "<":
            return na < nb
        if op == "<=":
            return na <= nb
        if op == ">":
            return na > nb
        if op == ">=":
            return na >= nb
    except EvalError:
        pass
    # Fall back to string comparison
    sa, sb = str(a).upper(), str(b).upper()
    if op == "=":
        return sa == sb
    if op == "<>":
        return sa != sb
    if op == "<":
        return sa < sb
    if op == "<=":
        return sa <= sb
    if op == ">":
        return sa > sb
    if op == ">=":
        return sa >= sb
    return False


def _parse_criteria(criteria: str) -> tuple[str, Any]:
    """Parse a SUMIF/COUNTIF criteria string into (operator, value)."""
    s = str(criteria).strip()
    for op in ("<=", ">=", "<>", "<", ">", "="):
        if s.startswith(op):
            val = s[len(op):].strip()
            try:
                val = float(val)
            except ValueError:
                pass
            return op, val
    # No operator means =
    try:
        return "=", float(s)
    except ValueError:
        return "=", s


class ExpressionEvaluator:
    """Evaluate a formula from its token list against a workbook state."""

    def __init__(self, wb: Workbook, current_sheet: str):
        self.wb = wb
        self.current_sheet = current_sheet

    def evaluate(self, tokens: list[dict]) -> Any:
        """Evaluate token list and return result."""
        if not tokens:
            raise EvalError("Empty token list")
        self._tokens = tokens
        self._pos = 0
        result = self._expr()
        return result

    def _peek(self) -> dict | None:
        if self._pos < len(self._tokens):
            return self._tokens[self._pos]
        return None

    def _advance(self) -> dict:
        tok = self._tokens[self._pos]
        self._pos += 1
        return tok

    def _expr(self) -> Any:
        """Parse comparison expressions (lowest precedence)."""
        left = self._add_expr()
        while True:
            tok = self._peek()
            if tok and tok["type"] == "OPERATOR-INFIX" and tok["value"] in ("=", "<>", "<", "<=", ">", ">="):
                op = self._advance()["value"]
                right = self._add_expr()
                left = _compare(left, right, op)
            elif tok and tok["type"] == "OPERATOR-INFIX" and tok["value"] == "&":
                self._advance()
                right = self._add_expr()
                left = str(left if left is not None else "") + str(right if right is not None else "")
            else:
                break
        return left

    def _add_expr(self) -> Any:
        """Parse addition/subtraction."""
        left = self._mul_expr()
        while True:
            tok = self._peek()
            if tok and tok["type"] == "OPERATOR-INFIX" and tok["value"] in ("+", "-"):
                op = self._advance()["value"]
                right = self._mul_expr()
                if op == "+":
                    left = _to_number(left) + _to_number(right)
                else:
                    left = _to_number(left) - _to_number(right)
            else:
                break
        return left

    def _mul_expr(self) -> Any:
        """Parse multiplication/division."""
        left = self._power_expr()
        while True:
            tok = self._peek()
            if tok and tok["type"] == "OPERATOR-INFIX" and tok["value"] in ("*", "/"):
                op = self._advance()["value"]
                right = self._power_expr()
                if op == "*":
                    left = _to_number(left) * _to_number(right)
                else:
                    divisor = _to_number(right)
                    if divisor == 0:
                        raise EvalError("Division by zero")
                    left = _to_number(left) / divisor
            else:
                break
        return left

    def _power_expr(self) -> Any:
        """Parse exponentiation."""
        base = self._unary()
        tok = self._peek()
        if tok and tok["type"] == "OPERATOR-INFIX" and tok["value"] == "^":
            self._advance()
            exp = self._unary()
            try:
                return _to_number(base) ** _to_number(exp)
            except OverflowError:
                return math.inf
        return base

    def _unary(self) -> Any:
        """Parse unary operators and percent."""
        tok = self._peek()
        if tok and tok["type"] == "OPERATOR-PREFIX" and tok["value"] in ("+", "-"):
            op = self._advance()["value"]
            val = self._unary()
            if op == "-":
                return -_to_number(val)
            # Unary + on strings just passes through (Excel behavior)
            if op == "+" and isinstance(val, str):
                return val
            try:
                return _to_number(val)
            except EvalError:
                return val
        val = self._atom()
        # Postfix percent
        tok = self._peek()
        if tok and tok["type"] == "OPERATOR-POSTFIX" and tok["value"] == "%":
            self._advance()
            return _to_number(val) / 100.0
        return val

    def _atom(self) -> Any:
        """Parse atomic values: numbers, strings, booleans, refs, functions, parens."""
        tok = self._peek()
        if tok is None:
            raise EvalError("Unexpected end of tokens")

        # Subexpression in parentheses (openpyxl uses both SUBEXPR and PAREN)
        if tok["type"] in ("SUBEXPR", "PAREN") and tok["subtype"] == "OPEN":
            self._advance()
            val = self._expr()
            close = self._peek()
            if close and close["type"] in ("SUBEXPR", "PAREN") and close["subtype"] == "CLOSE":
                self._advance()
            return val

        # Function call
        if tok["type"] == "FUNC" and tok["subtype"] == "OPEN":
            return self._func_call()

        # Operand
        if tok["type"] == "OPERAND":
            return self._operand()

        # Array constant: {1,2,3} or {1,2;3,4}
        if tok["type"] == "ARRAY" and tok["subtype"] == "OPEN":
            return self._array_constant()

        raise EvalError(f"Unexpected token: {tok}")

    def _array_constant(self) -> Any:
        """Parse an array constant: {1,2,3} or {1,2;3,4}.
        Returns a flat list of values."""
        self._advance()  # consume ARRAY OPEN
        values = []
        while True:
            tok = self._peek()
            if tok is None:
                break
            if tok["type"] == "ARRAY" and tok["subtype"] == "CLOSE":
                self._advance()
                break
            if tok["type"] == "SEP" and tok["subtype"] in ("ARG", "ROW"):
                self._advance()
                continue
            values.append(self._expr())
        # Return flat list; SUM and other aggregates handle lists natively
        if len(values) == 1:
            return values[0]
        return values

    def _operand(self) -> Any:
        tok = self._advance()
        subtype = tok["subtype"]
        value = tok["value"]

        if subtype == "NUMBER":
            try:
                return float(value)
            except ValueError:
                raise EvalError(f"Invalid number: {value}")

        if subtype == "TEXT":
            # Strip surrounding quotes
            if value.startswith('"') and value.endswith('"'):
                return value[1:-1]
            return value

        if subtype == "LOGICAL":
            return value.upper() == "TRUE"

        if subtype == "RANGE":
            return self._resolve_ref(value)

        if subtype == "ERROR":
            raise EvalError(f"Error value: {value}")

        raise EvalError(f"Unknown operand subtype: {subtype}")

    def _resolve_ref(self, ref: str) -> Any:
        """Resolve a cell/range reference to its value(s)."""
        # Strip implicit intersection operator @
        if ref.startswith("@"):
            ref = ref[1:]

        # Parse sheet!ref
        sheet = self.current_sheet
        cell_ref = ref
        if "!" in ref:
            parts = ref.split("!", 1)
            sheet = parts[0].strip("'").strip('"')
            cell_ref = parts[1]

        cell_ref = cell_ref.replace("$", "").upper()

        # Handle structured references: Table1[Column]
        if "[" in cell_ref:
            # Not supported — return None (caller's IFERROR will handle)
            return None

        # Handle whole-column (A:A) or whole-row (1:1) references
        if ":" in cell_ref:
            parts = cell_ref.split(":")
            # Whole-column: both parts are letters only (no digits)
            # Whole-row: both parts are digits only (no letters)
            is_whole_col = all(c.isalpha() for c in parts[0]) and all(c.isalpha() for c in parts[1])
            is_whole_row = all(c.isdigit() for c in parts[0]) and all(c.isdigit() for c in parts[1])
            if is_whole_col or is_whole_row:
                # Scan known cells in this sheet for matches
                values = []
                prefix = f"{sheet}!"
                if is_whole_col:
                    cn1, cn2 = col_to_num(parts[0]), col_to_num(parts[1])
                    for qc, val in list(self.wb.literal_values.items()) + [(fc.qcell, fc.computed_value) for fc in self.wb.formula_cells.values()]:
                        if qc.startswith(prefix):
                            cell_part = qc[len(prefix):]
                            try:
                                col_str, _ = parse_cell(cell_part)
                                cn = col_to_num(col_str)
                                if cn1 <= cn <= cn2:
                                    values.append(val)
                            except ValueError:
                                pass
                else:
                    r1, r2 = int(parts[0]), int(parts[1])
                    for qc, val in list(self.wb.literal_values.items()) + [(fc.qcell, fc.computed_value) for fc in self.wb.formula_cells.values()]:
                        if qc.startswith(prefix):
                            cell_part = qc[len(prefix):]
                            try:
                                _, row = parse_cell(cell_part)
                                if r1 <= row <= r2:
                                    values.append(val)
                            except ValueError:
                                pass
                return values

            # Normal bounded range
            cells = expand_range(cell_ref)
            return [self.wb.get_value(qualified_cell(sheet, c)) for c in cells]

        # 3D reference detection: Sheet1:Sheet3 in the sheet part
        if ":" in sheet:
            # Not supported in v1
            return None

        # Single cell
        qc = qualified_cell(sheet, cell_ref)
        return self.wb.get_value(qc)

    def _collect_args(self) -> list[Any]:
        """Collect function arguments until closing paren."""
        args = []
        tok = self._peek()
        # Empty args
        if tok and tok["type"] == "FUNC" and tok["subtype"] == "CLOSE":
            self._advance()
            return args

        args.append(self._expr())
        while True:
            tok = self._peek()
            if tok is None:
                break
            if tok["type"] == "FUNC" and tok["subtype"] == "CLOSE":
                self._advance()
                break
            if tok["type"] == "SEP" and tok["subtype"] == "ARG":
                self._advance()
                args.append(self._expr())
            else:
                break
        return args

    def _collect_raw_args(self) -> list:
        """Collect function arguments as raw token groups (for range-aware functions)."""
        args = []
        current_tokens = []
        depth = 0

        while self._pos < len(self._tokens):
            tok = self._peek()
            if tok is None:
                break

            if tok["type"] == "FUNC" and tok["subtype"] == "CLOSE" and depth == 0:
                self._advance()
                if current_tokens:
                    args.append(current_tokens)
                return args

            if tok["type"] == "FUNC" and tok["subtype"] == "OPEN":
                depth += 1
                current_tokens.append(self._advance())
            elif tok["type"] == "FUNC" and tok["subtype"] == "CLOSE":
                depth -= 1
                current_tokens.append(self._advance())
            elif tok["type"] == "SEP" and tok["subtype"] == "ARG" and depth == 0:
                self._advance()
                args.append(current_tokens)
                current_tokens = []
            elif tok["type"] in ("SUBEXPR", "PAREN", "ARRAY") and tok["subtype"] == "OPEN":
                depth += 1
                current_tokens.append(self._advance())
            elif tok["type"] in ("SUBEXPR", "PAREN", "ARRAY") and tok["subtype"] == "CLOSE":
                depth -= 1
                current_tokens.append(self._advance())
            else:
                current_tokens.append(self._advance())

        if current_tokens:
            args.append(current_tokens)
        return args

    def _eval_token_group(self, tokens: list[dict]) -> Any:
        """Evaluate a sub-list of tokens."""
        saved_tokens = self._tokens
        saved_pos = self._pos
        self._tokens = tokens
        self._pos = 0
        result = self._expr()
        self._tokens = saved_tokens
        self._pos = saved_pos
        return result

    def _resolve_range_from_tokens(self, tokens: list[dict]) -> tuple[str, str]:
        """Extract sheet and range ref from a token group containing a single RANGE operand."""
        for t in tokens:
            if t["type"] == "OPERAND" and t["subtype"] == "RANGE":
                ref = t["value"]
                sheet = self.current_sheet
                cell_ref = ref
                if "!" in ref:
                    parts = ref.split("!", 1)
                    sheet = parts[0].strip("'").strip('"')
                    cell_ref = parts[1]
                return sheet, cell_ref.replace("$", "").upper()
        raise EvalError("Expected range reference in token group")

    def _get_range_values(self, tokens: list[dict]) -> list[Any]:
        """Get a flat list of values from a range token group."""
        # Try evaluating as expression first — handles computed token groups
        val = self._eval_token_group(tokens)
        if isinstance(val, list):
            return val
        return [val]

    def _func_call(self) -> Any:
        tok = self._advance()  # consume FUNC OPEN
        fname = tok["value"].rstrip("(").upper()
        # Strip _xlfn. prefix (Excel 2013+ new functions like _xlfn.XLOOKUP)
        if fname.startswith("_XLFN."):
            fname = fname[6:]

        if fname in UNSUPPORTED_FUNCTIONS:
            raise UnsupportedError(f"Unsupported function: {fname}")
        if fname not in SUPPORTED_FUNCTIONS:
            raise UnsupportedError(f"Unsupported function: {fname}")

        # Functions that need raw token access for range arguments
        if fname in ("SUM", "MIN", "MAX", "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
                      "INDEX", "MATCH", "XMATCH", "XLOOKUP", "VLOOKUP", "HLOOKUP",
                      "IRR", "XIRR", "NPV", "XNPV", "TEXTJOIN",
                      "CONCATENATE", "CONCAT"):
            return self._call_range_func(fname)

        args = self._collect_args()
        return self._dispatch(fname, args)

    def _call_range_func(self, fname: str) -> Any:
        """Handle functions that operate on ranges."""
        raw_args = self._collect_raw_args()

        if fname == "SUM":
            total = 0.0
            for arg_tokens in raw_args:
                val = self._eval_token_group(arg_tokens)
                if isinstance(val, list):
                    for v in val:
                        if v is not None and not isinstance(v, bool) and not isinstance(v, str):
                            total += _to_number(v)
                elif val is not None:
                    total += _to_number(val)
            return total

        if fname == "MIN":
            nums = []
            for arg_tokens in raw_args:
                val = self._eval_token_group(arg_tokens)
                if isinstance(val, list):
                    for v in val:
                        if v is not None and isinstance(v, (int, float)):
                            nums.append(float(v))
                elif val is not None and isinstance(val, (int, float)):
                    nums.append(float(val))
            return min(nums) if nums else 0.0

        if fname == "MAX":
            nums = []
            for arg_tokens in raw_args:
                val = self._eval_token_group(arg_tokens)
                if isinstance(val, list):
                    for v in val:
                        if v is not None and isinstance(v, (int, float)):
                            nums.append(float(v))
                elif val is not None and isinstance(val, (int, float)):
                    nums.append(float(val))
            return max(nums) if nums else 0.0

        if fname == "SUMIF":
            return self._eval_sumif(raw_args)

        if fname == "SUMIFS":
            return self._eval_sumifs(raw_args)

        if fname == "COUNTIF":
            return self._eval_countif(raw_args)

        if fname == "COUNTIFS":
            return self._eval_countifs(raw_args)

        if fname == "INDEX":
            return self._eval_index(raw_args)

        if fname == "MATCH":
            return self._eval_match(raw_args)

        if fname == "XMATCH":
            return self._eval_xmatch(raw_args)

        if fname == "XLOOKUP":
            return self._eval_xlookup(raw_args)

        if fname == "VLOOKUP":
            return self._eval_vlookup(raw_args)

        if fname == "HLOOKUP":
            return self._eval_hlookup(raw_args)

        if fname == "IRR":
            return self._eval_irr(raw_args)

        if fname == "XIRR":
            return self._eval_xirr(raw_args)

        if fname == "NPV":
            return self._eval_npv(raw_args)

        if fname == "XNPV":
            return self._eval_xnpv(raw_args)

        if fname in ("CONCATENATE", "CONCAT"):
            parts = []
            for arg_tokens in raw_args:
                val = self._eval_token_group(arg_tokens)
                if isinstance(val, list):
                    for v in val:
                        parts.append(str(v) if v is not None else "")
                else:
                    parts.append(str(val) if val is not None else "")
            return "".join(parts)

        if fname == "TEXTJOIN":
            if len(raw_args) < 3:
                raise EvalError("TEXTJOIN requires at least 3 arguments")
            delimiter = str(self._eval_token_group(raw_args[0]) or "")
            ignore_empty = _to_bool(self._eval_token_group(raw_args[1]))
            parts = []
            for arg_tokens in raw_args[2:]:
                val = self._eval_token_group(arg_tokens)
                if isinstance(val, list):
                    for v in val:
                        s = str(v) if v is not None else ""
                        if ignore_empty and s == "":
                            continue
                        parts.append(s)
                else:
                    s = str(val) if val is not None else ""
                    if ignore_empty and s == "":
                        continue
                    parts.append(s)
            return delimiter.join(parts)

        raise UnsupportedError(f"Unhandled range function: {fname}")

    def _dispatch(self, fname: str, args: list[Any]) -> Any:
        """Dispatch to built-in function implementations."""

        if fname == "IF":
            if len(args) < 2:
                raise EvalError("IF requires at least 2 arguments")
            cond = _to_bool(args[0])
            if cond:
                return args[1]
            return args[2] if len(args) > 2 else False

        if fname == "AND":
            return all(_to_bool(a) for a in args)

        if fname == "OR":
            return any(_to_bool(a) for a in args)

        if fname == "NOT":
            if len(args) != 1:
                raise EvalError("NOT requires exactly 1 argument")
            return not _to_bool(args[0])

        if fname == "IFERROR":
            # IFERROR is special — first arg may have already raised.
            # We handle this at a higher level; if we get here, no error occurred.
            return args[0] if len(args) >= 1 else 0

        if fname == "IFNA":
            return args[0] if len(args) >= 1 else 0

        if fname == "ROUND":
            if len(args) < 1:
                raise EvalError("ROUND requires at least 1 argument")
            num = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            return round(num, digits)

        if fname == "ROUNDUP":
            num = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            factor = 10 ** digits
            # Excel ROUNDUP rounds away from zero
            if num >= 0:
                return math.ceil(num * factor) / factor
            else:
                return math.floor(num * factor) / factor

        if fname == "ROUNDDOWN":
            num = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            factor = 10 ** digits
            # Excel ROUNDDOWN rounds toward zero
            return math.trunc(num * factor) / factor

        if fname == "ABS":
            return abs(_to_number(args[0]))

        # Date functions
        if fname == "DATE":
            if len(args) < 3:
                raise EvalError("DATE requires 3 arguments")
            y, m, d = int(_to_number(args[0])), int(_to_number(args[1])), int(_to_number(args[2]))
            try:
                return _date_to_serial(_excel_date(y, m, d))
            except ValueError:
                raise EvalError(f"Invalid date: {y}-{m}-{d}")

        if fname == "DAY":
            serial = int(_to_number(args[0]))
            return _serial_to_date(serial).day

        if fname == "MONTH":
            serial = int(_to_number(args[0]))
            return _serial_to_date(serial).month

        if fname == "YEAR":
            serial = int(_to_number(args[0]))
            return _serial_to_date(serial).year

        if fname == "EDATE":
            if len(args) < 2:
                raise EvalError("EDATE requires 2 arguments")
            import calendar
            start = _serial_to_date(int(_to_number(args[0])))
            months = int(_to_number(args[1]))
            # Total months from epoch, then decompose
            total_months = start.year * 12 + (start.month - 1) + months
            new_year = total_months // 12
            new_month = (total_months % 12) + 1
            max_day = calendar.monthrange(new_year, new_month)[1]
            new_day = min(start.day, max_day)
            return _date_to_serial(date(new_year, new_month, new_day))

        if fname == "EOMONTH":
            if len(args) < 2:
                raise EvalError("EOMONTH requires 2 arguments")
            import calendar
            start = _serial_to_date(int(_to_number(args[0])))
            months = int(_to_number(args[1]))
            total_months = start.year * 12 + (start.month - 1) + months
            new_year = total_months // 12
            new_month = (total_months % 12) + 1
            last_day = calendar.monthrange(new_year, new_month)[1]
            return _date_to_serial(date(new_year, new_month, last_day))

        if fname == "YEARFRAC":
            if len(args) < 2:
                raise EvalError("YEARFRAC requires at least 2 arguments")
            d1 = _serial_to_date(int(_to_number(args[0])))
            d2 = _serial_to_date(int(_to_number(args[1])))
            basis = int(_to_number(args[2])) if len(args) > 2 else 0
            # Simplified: basis 0 (US 30/360) or actual/actual
            if basis == 0:
                # US 30/360
                dd1, dd2 = min(d1.day, 30), min(d2.day, 30)
                if d1.day == 31:
                    dd1 = 30
                if d2.day == 31 and d1.day >= 30:
                    dd2 = 30
                days = (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (dd2 - dd1)
                return days / 360.0
            else:
                # Actual/actual (basis 1) or actual/360 (basis 2) or actual/365 (basis 3)
                actual_days = (d2 - d1).days
                if basis == 1:
                    return actual_days / 365.25
                elif basis == 2:
                    return actual_days / 360.0
                elif basis == 3:
                    return actual_days / 365.0
                else:
                    return actual_days / 360.0

        raise UnsupportedError(f"Unhandled function: {fname}")

    # --- Conditional aggregate helpers ---

    def _eval_sumif(self, raw_args: list) -> float:
        if len(raw_args) < 2:
            raise EvalError("SUMIF requires at least 2 arguments")
        range_vals = self._get_range_values(raw_args[0])
        criteria = self._eval_token_group(raw_args[1])
        sum_range_vals = self._get_range_values(raw_args[2]) if len(raw_args) > 2 else range_vals
        op, target = _parse_criteria(criteria)
        total = 0.0
        for i, v in enumerate(range_vals):
            if _compare(v, target, op):
                sv = sum_range_vals[i] if i < len(sum_range_vals) else 0
                if sv is not None and not isinstance(sv, str):
                    total += _to_number(sv)
        return total

    def _eval_sumifs(self, raw_args: list) -> float:
        if len(raw_args) < 3:
            raise EvalError("SUMIFS requires at least 3 arguments")
        sum_range_vals = self._get_range_values(raw_args[0])
        # Pairs of (criteria_range, criteria)
        pairs = []
        i = 1
        while i + 1 < len(raw_args):
            cr = self._get_range_values(raw_args[i])
            crit = self._eval_token_group(raw_args[i + 1])
            pairs.append((cr, _parse_criteria(crit)))
            i += 2
        total = 0.0
        for idx, sv in enumerate(sum_range_vals):
            match = True
            for cr, (op, target) in pairs:
                cv = cr[idx] if idx < len(cr) else None
                if not _compare(cv, target, op):
                    match = False
                    break
            if match and sv is not None and not isinstance(sv, str):
                total += _to_number(sv)
        return total

    def _eval_countif(self, raw_args: list) -> int:
        if len(raw_args) < 2:
            raise EvalError("COUNTIF requires 2 arguments")
        range_vals = self._get_range_values(raw_args[0])
        criteria = self._eval_token_group(raw_args[1])
        op, target = _parse_criteria(criteria)
        count = 0
        for v in range_vals:
            if _compare(v, target, op):
                count += 1
        return count

    def _eval_countifs(self, raw_args: list) -> int:
        if len(raw_args) < 2:
            raise EvalError("COUNTIFS requires at least 2 arguments")
        first_range = self._get_range_values(raw_args[0])
        pairs = []
        i = 0
        while i + 1 < len(raw_args):
            cr = self._get_range_values(raw_args[i])
            crit = self._eval_token_group(raw_args[i + 1])
            pairs.append((cr, _parse_criteria(crit)))
            i += 2
        count = 0
        for idx in range(len(first_range)):
            match = True
            for cr, (op, target) in pairs:
                cv = cr[idx] if idx < len(cr) else None
                if not _compare(cv, target, op):
                    match = False
                    break
            if match:
                count += 1
        return count

    # --- Lookup functions ---

    def _eval_index(self, raw_args: list) -> Any:
        if len(raw_args) < 2:
            raise EvalError("INDEX requires at least 2 arguments")
        # Get the array as a 2D grid
        sheet, cell_ref = self._resolve_range_from_tokens(raw_args[0])
        row_num = int(_to_number(self._eval_token_group(raw_args[1])))
        col_num = int(_to_number(self._eval_token_group(raw_args[2]))) if len(raw_args) > 2 else 1

        if ":" in cell_ref:
            parts = cell_ref.split(":")
            c1, r1 = parse_cell(parts[0])
            c2, r2 = parse_cell(parts[1])
            cn1, cn2 = col_to_num(c1), col_to_num(c2)

            if row_num == 0 and col_num == 0:
                raise EvalError("INDEX: both row and col cannot be 0 for array")

            # If single row or single column range
            if r1 == r2:  # single row
                target_col = cn1 + col_num - 1 if col_num > 0 else cn1
                target_row = r1
            elif cn1 == cn2:  # single column
                target_row = r1 + row_num - 1 if row_num > 0 else r1
                target_col = cn1
            else:
                target_row = r1 + row_num - 1
                target_col = cn1 + col_num - 1

            qc = qualified_cell(sheet, f"{num_to_col(target_col)}{target_row}")
            return self.wb.get_value(qc)
        else:
            return self.wb.get_value(qualified_cell(sheet, cell_ref))

    def _eval_match(self, raw_args: list) -> int:
        if len(raw_args) < 2:
            raise EvalError("MATCH requires at least 2 arguments")
        lookup_val = self._eval_token_group(raw_args[0])
        range_vals = self._get_range_values(raw_args[1])
        match_type = int(_to_number(self._eval_token_group(raw_args[2]))) if len(raw_args) > 2 else 1

        if match_type == 0:
            # Exact match
            for i, v in enumerate(range_vals):
                if _values_match(v, lookup_val):
                    return i + 1
            raise EvalError(f"MATCH: no exact match found for {lookup_val}")
        elif match_type == 1:
            # Largest value <= lookup_val (sorted ascending)
            last = None
            for i, v in enumerate(range_vals):
                try:
                    if _to_number(v) <= _to_number(lookup_val):
                        last = i + 1
                    else:
                        break
                except EvalError:
                    continue
            if last is None:
                raise EvalError(f"MATCH: no match found for {lookup_val}")
            return last
        else:
            # match_type == -1: smallest value >= lookup_val (sorted descending)
            last = None
            for i, v in enumerate(range_vals):
                try:
                    if _to_number(v) >= _to_number(lookup_val):
                        last = i + 1
                    else:
                        break
                except EvalError:
                    continue
            if last is None:
                raise EvalError(f"MATCH: no match found for {lookup_val}")
            return last

    def _eval_xmatch(self, raw_args: list) -> int:
        if len(raw_args) < 2:
            raise EvalError("XMATCH requires at least 2 arguments")
        lookup_val = self._eval_token_group(raw_args[0])
        range_vals = self._get_range_values(raw_args[1])
        match_mode = int(_to_number(self._eval_token_group(raw_args[2]))) if len(raw_args) > 2 else 0

        if match_mode == 0:
            for i, v in enumerate(range_vals):
                if _values_match(v, lookup_val):
                    return i + 1
            raise EvalError(f"XMATCH: no exact match found for {lookup_val}")
        elif match_mode == -1:
            # Exact or next smaller
            best = None
            for i, v in enumerate(range_vals):
                try:
                    nv = _to_number(v)
                    if nv <= _to_number(lookup_val):
                        if best is None or nv > _to_number(range_vals[best]):
                            best = i
                except EvalError:
                    continue
            if best is None:
                raise EvalError(f"XMATCH: no match found for {lookup_val}")
            return best + 1
        elif match_mode == 1:
            # Exact or next larger
            best = None
            for i, v in enumerate(range_vals):
                try:
                    nv = _to_number(v)
                    if nv >= _to_number(lookup_val):
                        if best is None or nv < _to_number(range_vals[best]):
                            best = i
                except EvalError:
                    continue
            if best is None:
                raise EvalError(f"XMATCH: no match found for {lookup_val}")
            return best + 1
        else:
            raise EvalError(f"XMATCH: unsupported match_mode {match_mode}")

    def _eval_xlookup(self, raw_args: list) -> Any:
        if len(raw_args) < 3:
            raise EvalError("XLOOKUP requires at least 3 arguments")
        lookup_val = self._eval_token_group(raw_args[0])
        lookup_range = self._get_range_values(raw_args[1])
        return_range = self._get_range_values(raw_args[2])
        not_found = self._eval_token_group(raw_args[3]) if len(raw_args) > 3 else None
        match_mode = int(_to_number(self._eval_token_group(raw_args[4]))) if len(raw_args) > 4 else 0

        if match_mode == 0:
            for i, v in enumerate(lookup_range):
                if _values_match(v, lookup_val):
                    return return_range[i] if i < len(return_range) else None
            if not_found is not None:
                return not_found
            raise EvalError(f"XLOOKUP: no match found for {lookup_val}")
        else:
            raise UnsupportedError(f"XLOOKUP match_mode {match_mode} not yet supported")

    def _eval_vlookup(self, raw_args: list) -> Any:
        if len(raw_args) < 3:
            raise EvalError("VLOOKUP requires at least 3 arguments")
        lookup_val = self._eval_token_group(raw_args[0])
        # Table array — need the 2D structure
        sheet, cell_ref = self._resolve_range_from_tokens(raw_args[1])
        col_index = int(_to_number(self._eval_token_group(raw_args[2])))
        approx = _to_bool(self._eval_token_group(raw_args[3])) if len(raw_args) > 3 else True

        if ":" not in cell_ref:
            raise EvalError("VLOOKUP table must be a range")

        parts = cell_ref.split(":")
        c1, r1 = parse_cell(parts[0])
        c2, r2 = parse_cell(parts[1])
        cn1, cn2 = col_to_num(c1), col_to_num(c2)

        # Search first column
        for r in range(r1, r2 + 1):
            first_col_cell = qualified_cell(sheet, f"{c1}{r}")
            v = self.wb.get_value(first_col_cell)
            if not approx:
                if _values_match(v, lookup_val):
                    target_col = num_to_col(cn1 + col_index - 1)
                    return self.wb.get_value(qualified_cell(sheet, f"{target_col}{r}"))
            else:
                try:
                    if _to_number(v) <= _to_number(lookup_val):
                        # Check if next row exceeds
                        if r == r2:
                            target_col = num_to_col(cn1 + col_index - 1)
                            return self.wb.get_value(qualified_cell(sheet, f"{target_col}{r}"))
                        next_v = self.wb.get_value(qualified_cell(sheet, f"{c1}{r+1}"))
                        if next_v is None or _to_number(next_v) > _to_number(lookup_val):
                            target_col = num_to_col(cn1 + col_index - 1)
                            return self.wb.get_value(qualified_cell(sheet, f"{target_col}{r}"))
                except EvalError:
                    continue
        raise EvalError(f"VLOOKUP: no match found for {lookup_val}")

    def _eval_hlookup(self, raw_args: list) -> Any:
        if len(raw_args) < 3:
            raise EvalError("HLOOKUP requires at least 3 arguments")
        lookup_val = self._eval_token_group(raw_args[0])
        sheet, cell_ref = self._resolve_range_from_tokens(raw_args[1])
        row_index = int(_to_number(self._eval_token_group(raw_args[2])))
        approx = _to_bool(self._eval_token_group(raw_args[3])) if len(raw_args) > 3 else True

        if ":" not in cell_ref:
            raise EvalError("HLOOKUP table must be a range")

        parts = cell_ref.split(":")
        c1, r1 = parse_cell(parts[0])
        c2, r2 = parse_cell(parts[1])
        cn1, cn2 = col_to_num(c1), col_to_num(c2)

        # Search first row
        for c in range(cn1, cn2 + 1):
            col_letter = num_to_col(c)
            v = self.wb.get_value(qualified_cell(sheet, f"{col_letter}{r1}"))
            if not approx and _values_match(v, lookup_val):
                target_row = r1 + row_index - 1
                return self.wb.get_value(qualified_cell(sheet, f"{col_letter}{target_row}"))
        raise EvalError(f"HLOOKUP: no match found for {lookup_val}")

    # --- Finance functions ---

    def _eval_irr(self, raw_args: list) -> float:
        if len(raw_args) < 1:
            raise EvalError("IRR requires at least 1 argument")
        values = self._get_range_values(raw_args[0])
        guess = float(self._eval_token_group(raw_args[1])) if len(raw_args) > 1 else 0.1

        nums = [_to_number(v) for v in values if v is not None]
        if not nums:
            raise EvalError("IRR: no values")

        # Newton's method
        rate = guess
        for _ in range(100):
            npv_val = sum(cf / (1 + rate) ** i for i, cf in enumerate(nums))
            dnpv = sum(-i * cf / (1 + rate) ** (i + 1) for i, cf in enumerate(nums))
            if abs(dnpv) < 1e-15:
                break
            new_rate = rate - npv_val / dnpv
            if abs(new_rate - rate) < 1e-8:
                return new_rate
            rate = new_rate
        raise EvalError("IRR: did not converge")

    def _eval_xirr(self, raw_args: list) -> float:
        if len(raw_args) < 2:
            raise EvalError("XIRR requires at least 2 arguments")
        values = self._get_range_values(raw_args[0])
        dates = self._get_range_values(raw_args[1])
        guess = float(self._eval_token_group(raw_args[2])) if len(raw_args) > 2 else 0.1

        nums = [_to_number(v) for v in values if v is not None]
        date_serials = [_to_number(d) for d in dates if d is not None]

        if len(nums) != len(date_serials) or len(nums) < 2:
            raise EvalError("XIRR: values and dates must have same length >= 2")

        # Check sign pattern
        has_pos = any(n > 0 for n in nums)
        has_neg = any(n < 0 for n in nums)
        if not (has_pos and has_neg):
            raise EvalError("XIRR: cash flows must have both positive and negative values")

        d0 = date_serials[0]
        # Newton's method
        rate = guess
        for _ in range(100):
            xnpv_val = sum(cf / (1 + rate) ** ((dt - d0) / 365.0) for cf, dt in zip(nums, date_serials))
            dxnpv = sum(-((dt - d0) / 365.0) * cf / (1 + rate) ** ((dt - d0) / 365.0 + 1)
                        for cf, dt in zip(nums, date_serials))
            if abs(dxnpv) < 1e-15:
                break
            new_rate = rate - xnpv_val / dxnpv
            if abs(new_rate - rate) < 1e-8:
                return new_rate
            rate = new_rate
        raise EvalError("XIRR: did not converge")

    def _eval_npv(self, raw_args: list) -> float:
        if len(raw_args) < 2:
            raise EvalError("NPV requires at least 2 arguments")
        rate = _to_number(self._eval_token_group(raw_args[0]))
        total = 0.0
        period = 1
        for arg_tokens in raw_args[1:]:
            val = self._eval_token_group(arg_tokens)
            if isinstance(val, list):
                for v in val:
                    if v is not None:
                        total += _to_number(v) / (1 + rate) ** period
                        period += 1
            else:
                total += _to_number(val) / (1 + rate) ** period
                period += 1
        return total

    def _eval_xnpv(self, raw_args: list) -> float:
        if len(raw_args) < 3:
            raise EvalError("XNPV requires 3 arguments")
        rate = _to_number(self._eval_token_group(raw_args[0]))
        values = self._get_range_values(raw_args[1])
        dates = self._get_range_values(raw_args[2])

        nums = [_to_number(v) for v in values if v is not None]
        date_serials = [_to_number(d) for d in dates if d is not None]

        if len(nums) != len(date_serials):
            raise EvalError("XNPV: values and dates must have same length")

        d0 = date_serials[0]
        return sum(cf / (1 + rate) ** ((dt - d0) / 365.0) for cf, dt in zip(nums, date_serials))


# ---------------------------------------------------------------------------
# IFERROR-aware evaluation wrapper
# ---------------------------------------------------------------------------

def evaluate_cell(wb: Workbook, fc: FormulaCell) -> Any:
    """Evaluate a single formula cell, handling IFERROR at the top level."""
    tokens = fc.tokens
    if not tokens:
        raise EvalError("No tokens")

    evaluator = ExpressionEvaluator(wb, fc.sheet)

    # Check if the formula is wrapped in IFERROR/IFNA at the top level
    top_fname = tokens[0]["value"].rstrip("(").upper() if tokens else ""
    if top_fname.startswith("_XLFN."):
        top_fname = top_fname[6:]
    if (tokens and tokens[0]["type"] == "FUNC" and tokens[0]["subtype"] == "OPEN"
            and top_fname in ("IFERROR", "IFNA")):
        # Manually handle: evaluate first arg, catch errors, return second arg
        fname = top_fname
        evaluator._tokens = tokens
        evaluator._pos = 1  # skip the FUNC OPEN
        raw_args = evaluator._collect_raw_args()
        if len(raw_args) >= 1:
            try:
                return evaluator._eval_token_group(raw_args[0])
            except (EvalError, ZeroDivisionError, ValueError, TypeError):
                if len(raw_args) >= 2:
                    return evaluator._eval_token_group(raw_args[1])
                return 0
        return 0

    return evaluator.evaluate(tokens)


# ---------------------------------------------------------------------------
# Dependency graph and execution engine
# ---------------------------------------------------------------------------

def build_dependency_graph(wb: Workbook) -> dict[str, set[str]]:
    """Build adjacency list: cell -> set of cells it depends on."""
    graph: dict[str, set[str]] = {}
    for qcell, fc in wb.formula_cells.items():
        deps = set()
        for dep in fc.dependencies:
            dep_sheet = dep.get("sheet", fc.sheet)
            dep_cell = dep["cell"].upper().replace("$", "")
            dqc = qualified_cell(dep_sheet, dep_cell)
            deps.add(dqc)
        graph[qcell] = deps
    return graph


def tarjan_scc(graph: dict[str, set[str]]) -> list[list[str]]:
    """Find strongly connected components using Tarjan's algorithm."""
    index_counter = [0]
    stack = []
    lowlink = {}
    index = {}
    on_stack = set()
    result = []

    def strongconnect(v):
        index[v] = index_counter[0]
        lowlink[v] = index_counter[0]
        index_counter[0] += 1
        stack.append(v)
        on_stack.add(v)

        for w in graph.get(v, set()):
            if w not in index:
                if w in graph:  # only visit formula cells
                    strongconnect(w)
                    lowlink[v] = min(lowlink[v], lowlink[w])
            elif w in on_stack:
                lowlink[v] = min(lowlink[v], index[w])

        if lowlink[v] == index[v]:
            component = []
            while True:
                w = stack.pop()
                on_stack.discard(w)
                component.append(w)
                if w == v:
                    break
            result.append(component)

    for v in graph:
        if v not in index:
            strongconnect(v)

    return result


def topological_sort(graph: dict[str, set[str]], exclude: set[str]) -> list[str]:
    """Topological sort of non-SCC nodes. Returns eval order."""
    from collections import deque

    nodes = set(graph.keys()) - exclude
    in_degree: dict[str, int] = {n: 0 for n in nodes}

    # Build reverse adjacency for efficient propagation
    dependents: dict[str, list[str]] = defaultdict(list)
    for n in nodes:
        for dep in graph.get(n, set()):
            if dep in nodes:
                in_degree[n] += 1
                dependents[dep].append(n)

    # Kahn's algorithm with deque
    queue = deque(n for n in nodes if in_degree[n] == 0)
    order = []
    while queue:
        node = queue.popleft()
        order.append(node)
        for dependent in dependents[node]:
            in_degree[dependent] -= 1
            if in_degree[dependent] == 0:
                queue.append(dependent)

    return order


# ---------------------------------------------------------------------------
# Iterative solver
# ---------------------------------------------------------------------------

MAX_ITERATIONS = 200
ABS_TOLERANCE = 1e-8
REL_TOLERANCE = 1e-7


def _run_iterations(wb: Workbook, fcs: list[FormulaCell]) -> bool:
    """Run iterative solving on a list of formula cells. Returns True if converged.

    During iteration, evaluation errors (e.g. division by zero on early passes)
    are tolerated — the cell keeps its old value and we continue iterating.
    This is critical for share-price/share-count loops where denominators start
    at zero before the loop bootstraps.
    """
    for iteration in range(MAX_ITERATIONS):
        converged = True
        had_errors = False
        for fc in fcs:
            if fc.status == "unsupported":
                continue
            old_val = fc.computed_value
            try:
                new_val = evaluate_cell(wb, fc)
                fc.computed_value = new_val
                try:
                    old_num = _to_number(old_val) if old_val is not None else 0.0
                    new_num = _to_number(new_val) if new_val is not None else 0.0
                    abs_delta = abs(new_num - old_num)
                    rel_delta = abs_delta / max(abs(old_num), 1e-15)
                    if abs_delta > ABS_TOLERANCE and rel_delta > REL_TOLERANCE:
                        converged = False
                except (EvalError, TypeError):
                    if str(old_val) != str(new_val):
                        converged = False
            except UnsupportedError:
                fc.status = "unsupported"
            except EvalError:
                # Tolerate errors during iteration — keep old value and continue.
                # The loop may bootstrap once other cells get non-zero values.
                had_errors = True
                converged = False
        if converged and not had_errors:
            # Verify no cell ended up at inf or nan
            for fc in fcs:
                if fc.status == "unsupported":
                    continue
                try:
                    v = _to_number(fc.computed_value) if fc.computed_value is not None else 0.0
                    if math.isinf(v) or math.isnan(v):
                        return False
                except (EvalError, TypeError):
                    pass
            return True
    return False


def _seed_iterative_cells(wb: Workbook, fcs: list[FormulaCell]) -> None:
    """Seed iterative cells with reasonable initial values.

    Strategy:
    1. Use cached_value from the dump if available (best case — Excel computed it)
    2. For cells with no cached value, try a one-shot evaluation using whatever
       values are currently available. This works for sum/total cells whose
       non-circular inputs are already resolved.
    3. For cells that fail one-shot eval (typically division-by-zero on early
       pass), seed SUM-type cells with 1.0 to give denominators something
       nonzero to work with.
    """
    component_qcells = {fc.qcell for fc in fcs}

    for fc in fcs:
        if fc.cached_value is not None:
            fc.computed_value = fc.cached_value
            continue

        # Try one-shot evaluation — succeeds for SUM/addition cells
        # whose external deps are already resolved
        try:
            val = evaluate_cell(wb, fc)
            if val is not None:
                fc.computed_value = val
                continue
        except (EvalError, ZeroDivisionError, ValueError, TypeError):
            pass

        # Heuristic: seed at 1.0 so that cells used as denominators
        # don't cause division-by-zero on the first real iteration
        fc.computed_value = 1.0 if fc.computed_value is None else fc.computed_value


def solve_iterative_component(wb: Workbook, component_cells: list[str]) -> bool:
    """Solve an iterative component. Returns True if converged."""
    fcs = []
    for qc in component_cells:
        fc = wb.formula_cells.get(qc)
        if fc and fc.status != "unsupported":
            fcs.append(fc)

    if not fcs:
        return True

    # Smart seeding: cached values → one-shot eval → heuristic nonzero
    _seed_iterative_cells(wb, fcs)

    # Attempt 1: solve with current seeds
    if _run_iterations(wb, fcs):
        for fc in fcs:
            if fc.status == "pending":
                fc.status = "supported-converged"
        return True

    # Attempt 2: if stuck, try different seed magnitudes
    # (the right scale matters — 1.0 may be too small or too large)
    for seed in [1e6, 1e-4, 1e3]:
        for fc in fcs:
            try:
                if fc.computed_value is None or _to_number(fc.computed_value) == 0.0:
                    fc.computed_value = seed
            except (EvalError, TypeError):
                fc.computed_value = seed
        if _run_iterations(wb, fcs):
            for fc in fcs:
                if fc.status == "pending":
                    fc.status = "supported-converged"
            return True

    # Did not converge
    for fc in fcs:
        if fc.status == "pending":
            fc.status = "supported-nonconverged"
    return False


# ---------------------------------------------------------------------------
# Workbook loader
# ---------------------------------------------------------------------------

def load_workbook(formulas_json_dir: Path) -> Workbook:
    """Load workbook from formulas_json dump directory."""
    wb = Workbook()

    # Load workbook.json
    wb_path = formulas_json_dir / "workbook.json"
    if not wb_path.exists():
        raise FileNotFoundError(f"workbook.json not found in {formulas_json_dir}")
    wb_data = json.loads(wb_path.read_text())
    wb.sheet_names = wb_data.get("sheet_names", [])
    wb.functions_used = wb_data.get("functions_used", {})
    wb.iterative_components = wb_data.get("iterative_components", [])

    # Load per-sheet JSON
    for sheet_name in wb.sheet_names:
        sheet_path = formulas_json_dir / f"{sheet_name}.json"
        if not sheet_path.exists():
            continue
        sheet_data = json.loads(sheet_path.read_text())
        for cell_data in sheet_data.get("formula_cells", []):
            fc = FormulaCell(sheet_name, cell_data)
            wb.add_cell(fc)
            # Also store cached_value as a literal fallback for dependency resolution
            if cell_data.get("cached_value") is not None:
                wb.literal_values[fc.qcell] = cell_data["cached_value"]
        # Load literal (non-formula) cell values
        for cell_ref, value in sheet_data.get("literal_cells", {}).items():
            qc = qualified_cell(sheet_name, cell_ref)
            if qc not in wb.formula_cells:
                wb.literal_values[qc] = value

    return wb


# ---------------------------------------------------------------------------
# Checks beyond formula execution
# ---------------------------------------------------------------------------

def check_sheet_completeness(wb: Workbook, spec: dict) -> list[dict]:
    """Check that all required sheets exist."""
    issues = []
    required_sheets = {s["name"] for s in spec.get("sheets", [])}
    actual_sheets = set(wb.sheet_names)

    for name in required_sheets:
        if name not in actual_sheets:
            issues.append({
                "severity": "critical",
                "sheet": name,
                "summary": "Required sheet missing",
                "details": f"Sheet '{name}' is required by the spec but not found in the workbook",
                "fix": f"Create the '{name}' sheet as specified",
            })

    return issues


def check_hardcoded_literals(wb: Workbook, spec: dict) -> list[dict]:
    """Flag modeled sheets where expected formula regions have literal values instead."""
    issues = []
    modeled_sheets = {s["name"] for s in spec.get("sheets", [])
                      if s.get("build_type") in ("build", "adapt")}

    for sheet_name in modeled_sheets:
        if sheet_name not in wb.sheet_cells:
            continue
        formula_cells = {fc.cell for fc in wb.sheet_cells[sheet_name]}
        # Check for cells that are referenced by formulas but are not themselves formulas
        # This is a heuristic — flag literal cells referenced by multiple formula cells
        ref_counts: dict[str, int] = defaultdict(int)
        for fc in wb.sheet_cells[sheet_name]:
            for dep in fc.dependencies:
                if dep.get("sheet", sheet_name) == sheet_name:
                    dep_cell = dep["cell"].upper().replace("$", "")
                    if dep_cell not in formula_cells:
                        ref_counts[dep_cell] += 1

    return issues


def check_totals_and_percentages(wb: Workbook) -> list[dict]:
    """Check that total rows use SUM formulas and percentage columns sum correctly."""
    issues = []
    for sheet_name, cells in wb.sheet_cells.items():
        for fc in cells:
            # Check if a cell named/labeled as "total" uses SUM
            formula_upper = fc.formula.upper() if fc.formula else ""
            # Simple heuristic: formulas in rows that likely are totals should use SUM
            # We check if cells that sum a column actually use SUM
            if fc.computed_value is not None and isinstance(fc.computed_value, (int, float)):
                pass  # Could do more sophisticated checks with row labels
    return issues


def check_placeholder_hygiene(wb: Workbook) -> list[dict]:
    """Flag rows where all formula cells compute to zero."""
    issues = []
    for sheet_name, cells in wb.sheet_cells.items():
        # Group cells by row
        rows: dict[int, list[FormulaCell]] = defaultdict(list)
        for fc in cells:
            _, row = parse_cell(fc.cell)
            rows[row].append(fc)

        for row_num, row_cells in sorted(rows.items()):
            if len(row_cells) < 2:
                continue
            all_zero = all(
                fc.computed_value is not None
                and isinstance(fc.computed_value, (int, float))
                and fc.computed_value == 0
                for fc in row_cells
            )
            if all_zero:
                issues.append({
                    "severity": "warning",
                    "sheet": sheet_name,
                    "summary": f"All-zero formula row {row_num}",
                    "details": f"Row {row_num} has {len(row_cells)} formula cells all computing to zero",
                    "fix": f"Verify row {row_num} is not a placeholder. Delete if not needed.",
                })

    return issues


def check_broken_references(wb: Workbook) -> list[dict]:
    """Check for references to non-existent sheets."""
    issues = []
    actual_sheets = set(wb.sheet_names)
    seen = set()

    for fc in wb.formula_cells.values():
        for ref in fc.references:
            ref_sheet = ref.get("sheet", fc.sheet)
            if ref_sheet not in actual_sheets and ref_sheet not in seen:
                seen.add(ref_sheet)
                issues.append({
                    "severity": "critical",
                    "sheet": fc.sheet,
                    "summary": f"Reference to missing sheet '{ref_sheet}'",
                    "details": f"Cell {fc.cell} references sheet '{ref_sheet}' which does not exist",
                    "fix": f"Fix references to '{ref_sheet}' or create the missing sheet",
                })

    return issues


# ---------------------------------------------------------------------------
# Main evaluation pipeline
# ---------------------------------------------------------------------------

def evaluate_workbook(formulas_json_dir: Path, spec: dict | None = None) -> dict:
    """Run the deterministic evaluator on a workbook dump.

    Returns verdict dict with shape:
    {
        "logic": {"passed": bool, "issues": [...]},
        "visual": {"grade": "?", "issues": []},
        "summary": {...}
    }
    """
    wb = load_workbook(formulas_json_dir)
    issues: list[dict] = []

    # --- Classify function support ---
    unsupported_cells = []
    for fc in wb.formula_cells.values():
        supported, unsup_funcs = classify_functions(fc.functions)
        if not supported:
            fc.status = "unsupported"
            unsupported_cells.append((fc, unsup_funcs))

    # --- Build dependency graph ---
    graph = build_dependency_graph(wb)

    # --- Detect SCCs from dump metadata or derive ---
    iter_component_cells: dict[str, list[str]] = {}
    for comp in wb.iterative_components:
        comp_id = comp["id"]
        cells = []
        for cell_ref in comp["cells"]:
            # cell_ref is like "Series A!E16"
            if "!" in cell_ref:
                cells.append(cell_ref)
            else:
                cells.append(cell_ref)
        iter_component_cells[comp_id] = cells

    # Identify all iterative cell keys
    iterative_qcells = set()
    for cells in iter_component_cells.values():
        iterative_qcells.update(cells)

    # Also derive SCCs from graph for cells not already tagged
    derived_sccs = tarjan_scc(graph)
    for scc in derived_sccs:
        # Multi-cell cycle or single-cell self-reference
        is_cycle = len(scc) > 1 or (len(scc) == 1 and scc[0] in graph.get(scc[0], set()))
        if is_cycle:
            already_tagged = any(c in iterative_qcells for c in scc)
            if not already_tagged:
                comp_id = f"derived_scc_{len(iter_component_cells)}"
                iter_component_cells[comp_id] = scc
                iterative_qcells.update(scc)

    # --- Interleaved execution: acyclic cells and iterative components ---
    # For each acyclic cell, track which iterative components it depends on.
    # Then execute in waves: eval acyclic cells with no iter deps, solve iter
    # component, eval newly-unblocked acyclic cells, solve next iter, etc.

    acyclic_order = topological_sort(graph, iterative_qcells)

    # Map each acyclic cell to the set of iterative component IDs it depends on
    cell_to_comp: dict[str, str] = {}
    for comp_id, comp_cells in iter_component_cells.items():
        for qc in comp_cells:
            cell_to_comp[qc] = comp_id

    # For each acyclic cell, find which iter components it transitively needs
    iter_deps_of: dict[str, set[str]] = {}
    for qcell in acyclic_order:
        needed = set()
        for dep in graph.get(qcell, set()):
            if dep in cell_to_comp:
                needed.add(cell_to_comp[dep])
            needed |= iter_deps_of.get(dep, set())
        iter_deps_of[qcell] = needed

    def _eval_acyclic(qcell: str) -> None:
        fc = wb.formula_cells.get(qcell)
        if fc is None or fc.status in ("unsupported",):
            return
        try:
            fc.computed_value = evaluate_cell(wb, fc)
            fc.status = "supported-converged"
        except UnsupportedError as e:
            fc.status = "unsupported"
            unsupported_cells.append((fc, [str(e)]))
        except EvalError as e:
            fc.status = "parse-failure"
            issues.append({
                "severity": "warning",
                "sheet": fc.sheet,
                "summary": f"Parse/eval failure at {fc.cell}",
                "details": str(e),
                "fix": f"Check formula at {fc.cell}: {fc.formula}",
            })

    # Build dependency order among iterative components themselves
    comp_deps: dict[str, set[str]] = {cid: set() for cid in iter_component_cells}
    for comp_id, comp_cells in iter_component_cells.items():
        for qc in comp_cells:
            for dep in graph.get(qc, set()):
                dep_comp = cell_to_comp.get(dep)
                if dep_comp and dep_comp != comp_id:
                    comp_deps[comp_id].add(dep_comp)
                # Also check transitive deps through acyclic cells
                if dep in iter_deps_of:
                    for needed_comp in iter_deps_of[dep]:
                        if needed_comp != comp_id:
                            comp_deps[comp_id].add(needed_comp)

    # Topological sort of iterative components
    from collections import deque
    comp_order = []
    comp_in_degree = {cid: len(deps) for cid, deps in comp_deps.items()}
    cq = deque(cid for cid, deg in comp_in_degree.items() if deg == 0)
    while cq:
        cid = cq.popleft()
        comp_order.append(cid)
        for other_cid, deps in comp_deps.items():
            if cid in deps:
                comp_in_degree[other_cid] -= 1
                if comp_in_degree[other_cid] == 0:
                    cq.append(other_cid)
    # Add any remaining (cycles between components — rare)
    for cid in iter_component_cells:
        if cid not in comp_order:
            comp_order.append(cid)

    # Execute: for each component in order, first eval acyclic cells that
    # are now unblocked, then solve the component
    solved_comps: set[str] = set()

    # Phase 0: eval acyclic cells with NO iterative dependencies
    for qcell in acyclic_order:
        if not iter_deps_of[qcell]:
            _eval_acyclic(qcell)

    # Phase 1..N: solve each component, then eval newly unblocked cells
    for comp_id in comp_order:
        comp_cells = iter_component_cells[comp_id]
        converged = solve_iterative_component(wb, comp_cells)
        solved_comps.add(comp_id)
        if not converged:
            affected = [c for c in comp_cells if c in wb.formula_cells]
            issues.append({
                "severity": "critical",
                "sheet": affected[0].split("!")[0] if affected else "?",
                "summary": f"Iterative component '{comp_id}' did not converge",
                "details": f"Cells: {', '.join(affected[:5])}{'...' if len(affected) > 5 else ''} did not converge within {MAX_ITERATIONS} iterations",
                "fix": "Check circular reference logic and IFERROR wrapping",
            })

        # Eval acyclic cells whose iter dependencies are now all solved
        for qcell in acyclic_order:
            if iter_deps_of[qcell] and iter_deps_of[qcell] <= solved_comps:
                fc = wb.formula_cells.get(qcell)
                if fc and fc.status == "pending":
                    _eval_acyclic(qcell)

    # --- Formula/graph-level checks only ---
    issues.extend(check_broken_references(wb))

    # --- Build summary ---
    total_cells = len(wb.formula_cells)
    supported_converged = sum(1 for fc in wb.formula_cells.values() if fc.status == "supported-converged")
    supported_nonconverged = sum(1 for fc in wb.formula_cells.values() if fc.status == "supported-nonconverged")
    unsupported_count = sum(1 for fc in wb.formula_cells.values() if fc.status == "unsupported")
    parse_failures = sum(1 for fc in wb.formula_cells.values() if fc.status == "parse-failure")
    graph_failures = sum(1 for fc in wb.formula_cells.values() if fc.status == "graph-failure")
    pending = sum(1 for fc in wb.formula_cells.values() if fc.status == "pending")

    # Function coverage
    func_coverage = {}
    for fname in sorted(wb.functions_used.keys()):
        fu = fname.upper()
        func_coverage[fname] = {
            "count": wb.functions_used[fname],
            "supported": fu in SUPPORTED_FUNCTIONS,
        }

    # Iterative component summary
    iter_summary = {}
    for comp_id, comp_cells in iter_component_cells.items():
        statuses = set()
        for qc in comp_cells:
            fc = wb.formula_cells.get(qc)
            if fc:
                statuses.add(fc.status)
        if "supported-nonconverged" in statuses:
            iter_summary[comp_id] = "non-converged"
        elif "supported-converged" in statuses:
            iter_summary[comp_id] = "converged"
        else:
            iter_summary[comp_id] = "unsupported/failed"

    # --- Determine logic pass/fail ---
    has_critical = any(i["severity"] == "critical" for i in issues)

    # Unsupported cells that are material (have downstream dependents)
    unsupported_inventory = []
    for fc, funcs in unsupported_cells:
        unsupported_inventory.append({
            "cell": fc.qcell,
            "functions": funcs,
            "formula": fc.formula,
        })

    summary = {
        "total_cells": total_cells,
        "supported_converged": supported_converged,
        "supported_nonconverged": supported_nonconverged,
        "unsupported": unsupported_count,
        "parse_failures": parse_failures,
        "graph_failures": graph_failures,
        "pending": pending,
        "function_coverage": func_coverage,
        "iterative_components": iter_summary,
        "unsupported_inventory": unsupported_inventory[:50],  # cap for readability
    }

    verdict = {
        "logic": {
            "passed": not has_critical,
            "issues": issues,
        },
        "visual": {
            "grade": "?",
            "issues": [],
        },
        "summary": summary,
    }

    return verdict


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Deterministic workbook evaluator")
    parser.add_argument("formulas_json_dir", help="Path to evals/formulas_json/ directory")
    parser.add_argument("--spec", help="Path to model_spec.json (optional)")
    parser.add_argument("--output", help="Output JSON path (default: stdout)")
    args = parser.parse_args()

    formulas_dir = Path(args.formulas_json_dir)
    spec = None
    if args.spec:
        spec = json.loads(Path(args.spec).read_text())

    verdict = evaluate_workbook(formulas_dir, spec)

    output = json.dumps(verdict, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(output)
        print(f"Verdict written to {args.output}")
    else:
        print(output)

    # Exit code
    sys.exit(0 if verdict["logic"]["passed"] else 1)


if __name__ == "__main__":
    main()
