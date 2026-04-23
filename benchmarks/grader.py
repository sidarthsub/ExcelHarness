"""Rubric-driven grader for candidate .xlsx files.

Reads grading.yaml, executes each check against the candidate workbook,
and returns a structured result:

    {
      "accuracy": float in [0, 1],
      "passed": int,
      "total": int,
      "weighted_score": float,
      "total_weight": float,
      "checks": [
        {"id": ..., "type": ..., "weight": ..., "passed": bool,
         "actual": ..., "expected": ..., "error": None | str},
        ...
      ]
    }

Negative-weight checks (e.g. hardcode_penalty) subtract from the
weighted_score when they trigger. Accuracy is max(0, weighted / sum_of_positive_weights).
"""
from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
import yaml

from benchmarks.recalc import recalc_xlsx


# ---- check dispatch --------------------------------------------------------


@dataclass
class CandidateWorkbook:
    path: Path
    wb_values: openpyxl.Workbook   # data_only=True — computed values
    wb_formulas: openpyxl.Workbook  # data_only=False — formula strings

    @classmethod
    def load(cls, path: Path) -> "CandidateWorkbook":
        return cls(
            path=path,
            wb_values=openpyxl.load_workbook(path, data_only=True),
            wb_formulas=openpyxl.load_workbook(path, data_only=False),
        )

    def sheet_names(self) -> list[str]:
        return list(self.wb_values.sheetnames)

    def cell_value(self, sheet: str, cell: str) -> Any:
        if sheet not in self.wb_values.sheetnames:
            raise KeyError(f"sheet '{sheet}' not in workbook (have: {self.sheet_names()})")
        return self.wb_values[sheet][cell].value

    def cell_formula(self, sheet: str, cell: str) -> Any:
        return self.wb_formulas[sheet][cell].value


def _tol_match(actual: float, expected: float, tol_abs: float | None, tol_rel: float | None) -> bool:
    if tol_abs is None and tol_rel is None:
        tol_rel = 1e-4
    if tol_abs is not None and abs(actual - expected) <= tol_abs:
        return True
    if tol_rel is not None and expected != 0 and abs(actual - expected) / abs(expected) <= tol_rel:
        return True
    if tol_rel is not None and expected == 0 and abs(actual) <= (tol_abs or 1e-9):
        return True
    return False


def _coerce_number(x: Any) -> float | None:
    if x is None:
        return None
    if isinstance(x, bool):
        return None
    if isinstance(x, (int, float)):
        if isinstance(x, float) and math.isnan(x):
            return None
        return float(x)
    if isinstance(x, str):
        s = x.strip().replace(",", "").replace("$", "").replace("%", "")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _iter_range_cells(sheet, addr: str):
    for row in sheet[addr]:
        for cell in row:
            yield cell


def check_numeric(cand: CandidateWorkbook, spec: dict) -> dict:
    sheet = spec["sheet"]
    cell = spec["cell"]
    actual_raw = cand.cell_value(sheet, cell)
    actual = _coerce_number(actual_raw)
    expected = float(spec["expected"])
    tol_abs = spec.get("tolerance_abs")
    tol_rel = spec.get("tolerance_rel")
    if tol_abs is not None:
        tol_abs = float(tol_abs)
    if tol_rel is not None:
        tol_rel = float(tol_rel)
    if actual is None:
        return {"passed": False, "actual": actual_raw, "expected": expected,
                "error": f"cell {sheet}!{cell} is not numeric (got {actual_raw!r})"}
    return {"passed": _tol_match(actual, expected, tol_abs, tol_rel),
            "actual": actual, "expected": expected, "error": None}


def check_sheet_exists(cand: CandidateWorkbook, spec: dict) -> dict:
    target = spec["sheet"]
    found = target in cand.sheet_names()
    return {"passed": found, "actual": cand.sheet_names(), "expected": target,
            "error": None if found else f"sheet '{target}' not present"}


def check_label_exists(cand: CandidateWorkbook, spec: dict) -> dict:
    needle = spec["text"]
    scope = spec.get("scope", "sheet")
    if scope == "workbook":
        sheets = cand.sheet_names()
    else:
        sheets = [spec["sheet"]]
    needle_l = needle.lower()
    for s in sheets:
        if s not in cand.wb_values.sheetnames:
            continue
        ws = cand.wb_values[s]
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str) and needle_l in v.lower():
                    return {"passed": True, "actual": v, "expected": needle, "error": None}
    return {"passed": False, "actual": None, "expected": needle,
            "error": f"label '{needle}' not found in {scope}"}


def check_formula_driven(cand: CandidateWorkbook, spec: dict) -> dict:
    sheet = spec["sheet"]
    cell = spec["cell"]
    raw = cand.cell_formula(sheet, cell)
    is_formula = isinstance(raw, str) and raw.startswith("=")
    return {"passed": is_formula, "actual": raw, "expected": "formula",
            "error": None if is_formula else f"{sheet}!{cell} is a literal, not a formula"}


def check_bs_balances(cand: CandidateWorkbook, spec: dict) -> dict:
    sheet = spec["sheet"]
    assets_row = spec["assets_row"]
    le_row = spec["liab_equity_row"]
    columns = spec["columns"]  # e.g. ["B", "C", "D", "E", "F"]
    tol_abs = spec.get("tolerance_abs", 1.0)
    tol_rel = spec.get("tolerance_rel", 1e-4)
    errs = []
    for col in columns:
        a = _coerce_number(cand.cell_value(sheet, f"{col}{assets_row}"))
        l = _coerce_number(cand.cell_value(sheet, f"{col}{le_row}"))
        if a is None or l is None:
            errs.append(f"col {col}: missing values A={a} L+E={l}")
            continue
        if not _tol_match(a, l, tol_abs, tol_rel):
            errs.append(f"col {col}: Assets={a:.2f} ≠ L+E={l:.2f}")
    return {"passed": not errs, "actual": errs or "balanced", "expected": "A = L+E",
            "error": "; ".join(errs) if errs else None}


def check_hardcode_penalty(cand: CandidateWorkbook, spec: dict) -> dict:
    """Penalize if scan_sheet contains the literal value of source_cell anywhere
    as a hardcoded number instead of a formula reference.

    source_cell format: "inputs/file.xlsx!Sheet!Cell" (relative to task dir if bare)
    """
    source = spec["source_cell"]
    scan_sheet = spec["sheet"]
    # source can be:  "<path>!<sheet>!<cell>"  or  "<sheet>!<cell>" (same workbook)
    parts = source.split("!")
    if len(parts) == 3:
        src_path, src_sheet, src_cell = parts
        task_dir = spec.get("_task_dir")
        abs_path = Path(src_path) if Path(src_path).is_absolute() else Path(task_dir) / src_path
        src_wb = openpyxl.load_workbook(abs_path, data_only=True)
        val = src_wb[src_sheet][src_cell].value
        src_wb.close()
    else:
        return {"passed": False, "actual": None, "expected": None,
                "error": f"source_cell must be path!sheet!cell, got {source}"}
    src_num = _coerce_number(val)
    if src_num is None:
        return {"passed": True, "actual": None, "expected": None,
                "error": f"source value {val!r} not numeric — penalty skipped"}
    # Scan scan_sheet for a literal (non-formula) cell with the same value.
    ws_vals = cand.wb_values[scan_sheet]
    ws_forms = cand.wb_formulas[scan_sheet]
    hits = []
    for row in ws_forms.iter_rows():
        for cell in row:
            f = cell.value
            if isinstance(f, str) and f.startswith("="):
                continue  # formula — not hardcoded
            v = ws_vals[cell.coordinate].value
            vn = _coerce_number(v)
            if vn is not None and _tol_match(vn, src_num, 1e-6, 1e-6):
                hits.append(f"{cell.coordinate}={v}")
    # passed = no hits (no penalty triggered)
    return {"passed": not hits, "actual": hits, "expected": "no hardcoded source values",
            "error": f"literal match(es) for {src_num}: {hits}" if hits else None}


def check_value_range(cand: CandidateWorkbook, spec: dict) -> dict:
    sheet = spec["sheet"]
    cell = spec["cell"]
    raw = cand.cell_value(sheet, cell)
    v = _coerce_number(raw)
    lo = spec.get("min", float("-inf"))
    hi = spec.get("max", float("inf"))
    if v is None:
        return {"passed": False, "actual": raw, "expected": f"[{lo}, {hi}]",
                "error": f"cell {sheet}!{cell} is not numeric"}
    ok = lo <= v <= hi
    return {"passed": ok, "actual": v, "expected": f"[{lo}, {hi}]",
            "error": None if ok else f"{v} outside [{lo}, {hi}]"}


def check_range_sum(cand: CandidateWorkbook, spec: dict) -> dict:
    sheet = spec["sheet"]
    rng = spec["range"]
    tol_abs = spec.get("tolerance_abs")
    tol_rel = spec.get("tolerance_rel")
    expected = float(spec["expected"])
    ws = cand.wb_values[sheet]
    total = 0.0
    for cell in _iter_range_cells(ws, rng):
        v = _coerce_number(cell.value)
        if v is not None:
            total += v
    return {"passed": _tol_match(total, expected, tol_abs, tol_rel),
            "actual": total, "expected": expected, "error": None}


def check_llm_judge(cand: CandidateWorkbook, spec: dict) -> dict:
    return {"passed": True, "actual": None, "expected": None,
            "error": "llm_judge not implemented — auto-pass stub"}


import re as _re

_TOKEN_SPLIT = _re.compile(r"[^a-z0-9%]+")


def _tokens(text: str) -> list[str]:
    """Lowercase tokenization, split on non-alphanumeric (keeps %)."""
    return [t for t in _TOKEN_SPLIT.split(text.lower()) if t]


def _label_alias_matches(alias: str, cell_text: str) -> bool:
    """Token-AND match: every non-trivial token in the alias must appear
    as a substring of cell_text (case-insensitive). Single-word aliases
    collapse to simple substring match.
    """
    alias_tokens = _tokens(alias)
    if not alias_tokens:
        return False
    cell_l = cell_text.lower()
    return all(tok in cell_l for tok in alias_tokens)


def _find_label_cells(cand: CandidateWorkbook, sheet: str, labels: list[str]):
    """Find ALL cells on the sheet whose string value matches any alias.

    Returns a list of (row, col, cell_text, matched_alias) in row-major order.
    Use plural-find so check functions can pick the first hit that has the
    required neighbor (e.g. section headers without a numeric value are
    naturally skipped).
    """
    ws = cand.wb_values[sheet]
    hits: list[tuple[int, int, str, str]] = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str):
                continue
            for alias in labels:
                if alias and _label_alias_matches(alias, v):
                    hits.append((cell.row, cell.column, v, alias))
                    break  # one alias per cell is enough
    return hits


def _find_label_cell(cand: CandidateWorkbook, sheet: str, labels: list[str]):
    """Legacy single-hit API used for formula_by_label / label-only lookups.
    Prefers the first match but returns (row, col, text) or None.
    """
    hits = _find_label_cells(cand, sheet, labels)
    if not hits:
        return None
    r, c, text, _alias = hits[0]
    return r, c, text


def _label_value_map(cand: CandidateWorkbook, sheet: str) -> list[dict]:
    """Return every (label, adjacent_value) pair on a sheet, useful for
    diagnosing why a label lookup failed.
    """
    out = []
    ws = cand.wb_values[sheet]
    for row in ws.iter_rows():
        cells = [c for c in row if c.value not in (None, "")]
        if not cells:
            continue
        # Find the first string cell; report its value + each numeric/text to its right.
        label_cell = next((c for c in cells if isinstance(c.value, str)), None)
        if label_cell is None:
            continue
        # Collect ≤3 adjacent non-empty values to the right
        adj = []
        for other in cells:
            if other is label_cell:
                continue
            if other.column > label_cell.column:
                adj.append({"col": other.coordinate, "value": other.value})
        out.append({"row": label_cell.row,
                    "label": label_cell.value,
                    "adjacent": adj[:3]})
    return out


def _labels_from_spec(spec: dict) -> list[str]:
    """Accept either `label: "..."` (single) or `labels: [...]` (list)."""
    if "labels" in spec and isinstance(spec["labels"], list):
        return [str(l) for l in spec["labels"] if l]
    if "label" in spec:
        return [str(spec["label"])]
    return []


def check_numeric_by_label(cand: CandidateWorkbook, spec: dict) -> dict:
    """Find the row containing any of the label aliases; check the numeric
    value at (row + row_offset, col + col_offset). col_offset defaults to 1
    (cell to the right). Layout-drift-tolerant version of `numeric`.

    Accepts `label: "..."` OR `labels: ["alias1", "alias2", ...]`.

    If multiple cells match the aliases, picks the FIRST match whose
    target offset has a numeric value — so section headers (no adjacent
    number) naturally get skipped over.
    """
    sheet = spec["sheet"]
    labels = _labels_from_spec(spec)
    col_offset = int(spec.get("col_offset", 1))
    row_offset = int(spec.get("row_offset", 0))
    tol_abs = spec.get("tolerance_abs")
    tol_rel = spec.get("tolerance_rel")
    if tol_abs is not None: tol_abs = float(tol_abs)
    if tol_rel is not None: tol_rel = float(tol_rel)
    expected = float(spec["expected"])

    hits = _find_label_cells(cand, sheet, labels)
    if not hits:
        return {"passed": False, "actual": None, "expected": expected,
                "error": f"none of labels {labels} found on sheet '{sheet}'",
                "diagnostics": {"sheet_label_value_map": _label_value_map(cand, sheet)}}
    # Try each match in order; first one with a numeric neighbor wins.
    attempted = []
    for r, c, text, alias in hits:
        target = cand.wb_values[sheet].cell(row=r + row_offset, column=c + col_offset)
        actual = _coerce_number(target.value)
        attempted.append({"row": r, "label": text, "neighbor": target.coordinate, "value": target.value})
        if actual is None:
            continue
        return {"passed": _tol_match(actual, expected, tol_abs, tol_rel),
                "actual": actual, "expected": expected,
                "matched_label": text, "error": None}
    # Every hit had a non-numeric neighbor
    return {"passed": False, "actual": None, "expected": expected,
            "error": f"found {len(hits)} label match(es) but none had a numeric neighbor",
            "diagnostics": {"attempted": attempted,
                            "sheet_label_value_map": _label_value_map(cand, sheet)}}


def check_formula_by_label(cand: CandidateWorkbook, spec: dict) -> dict:
    """Cell at (label + offset) must contain a formula, not a literal.
    Scans all label matches; first match whose target is a formula wins.
    """
    sheet = spec["sheet"]
    labels = _labels_from_spec(spec)
    col_offset = int(spec.get("col_offset", 1))
    row_offset = int(spec.get("row_offset", 0))
    hits = _find_label_cells(cand, sheet, labels)
    if not hits:
        return {"passed": False, "actual": None, "expected": "formula",
                "error": f"none of labels {labels} found on sheet '{sheet}'",
                "diagnostics": {"sheet_label_value_map": _label_value_map(cand, sheet)}}
    attempted = []
    for r, c, text, alias in hits:
        target = cand.wb_formulas[sheet].cell(row=r + row_offset, column=c + col_offset)
        raw = target.value
        attempted.append({"row": r, "label": text, "neighbor": target.coordinate, "raw": raw})
        if isinstance(raw, str) and raw.startswith("="):
            return {"passed": True, "actual": raw, "expected": "formula",
                    "matched_label": text, "error": None}
    return {"passed": False, "actual": None, "expected": "formula",
            "error": f"found {len(hits)} label match(es) but none had a formula at the expected offset",
            "diagnostics": {"attempted": attempted}}


def _parse_cell_addr(addr: str) -> tuple[str, str]:
    """Parse 'Sheet!Cell' or \"'Sheet Name'!Cell\". Returns (sheet, cell)."""
    if "!" not in addr:
        raise ValueError(f"expected Sheet!Cell format, got {addr!r}")
    sheet_part, cell = addr.rsplit("!", 1)
    sheet = sheet_part.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1]
    return sheet, cell.strip()


def _load_results_json(candidate_xlsx: Path) -> dict | None:
    """Look for a results.json alongside the candidate xlsx."""
    results_path = candidate_xlsx.parent / "results.json"
    if not results_path.exists():
        return None
    try:
        return json.loads(results_path.read_text())
    except json.JSONDecodeError as e:
        raise RuntimeError(f"results.json parse failed: {e}")


def check_output_key(cand: CandidateWorkbook, spec: dict) -> dict:
    """Dereference spec.key via results.json, check the resolved cell's value.

    Flow:
      1. Read `results.json` (next to candidate xlsx) — caller injected as spec._results.
      2. Look up spec.key → get "Sheet!Cell" address.
      3. Read that cell's computed value from the candidate xlsx.
      4. Compare to spec.expected with numeric tolerance.

    Optional spec.require_formula=True also verifies the cell contains a formula.
    """
    key = spec["key"]
    expected = float(spec["expected"])
    tol_abs = spec.get("tolerance_abs")
    tol_rel = spec.get("tolerance_rel")
    if tol_abs is not None: tol_abs = float(tol_abs)
    if tol_rel is not None: tol_rel = float(tol_rel)
    require_formula = bool(spec.get("require_formula", False))

    results = spec.get("_results")
    if results is None:
        return {"passed": False, "actual": None, "expected": expected,
                "error": "results.json not found next to candidate xlsx"}
    if key not in results:
        return {"passed": False, "actual": None, "expected": expected,
                "error": f"key '{key}' missing from results.json",
                "diagnostics": {"available_keys": list(results.keys())}}
    addr = results[key]
    if not isinstance(addr, str):
        return {"passed": False, "actual": addr, "expected": expected,
                "error": f"results.json[{key!r}] should be a cell address string, got {type(addr).__name__}"}
    try:
        sheet, cell = _parse_cell_addr(addr)
    except ValueError as e:
        return {"passed": False, "actual": addr, "expected": expected,
                "error": str(e)}
    if sheet not in cand.wb_values.sheetnames:
        return {"passed": False, "actual": addr, "expected": expected,
                "error": f"sheet '{sheet}' (from results.json[{key!r}]) not in workbook"}
    try:
        raw_val = cand.wb_values[sheet][cell].value
    except Exception as e:
        return {"passed": False, "actual": addr, "expected": expected,
                "error": f"could not read {addr}: {e}"}
    actual = _coerce_number(raw_val)
    if actual is None:
        return {"passed": False, "actual": raw_val, "expected": expected,
                "error": f"{addr} is not numeric (got {raw_val!r})"}
    passed = _tol_match(actual, expected, tol_abs, tol_rel)
    if passed and require_formula:
        raw_formula = cand.wb_formulas[sheet][cell].value
        if not (isinstance(raw_formula, str) and raw_formula.startswith("=")):
            return {"passed": False, "actual": raw_val, "expected": expected,
                    "error": f"{addr} value matches but cell is a literal, require_formula=True",
                    "resolved": addr}
    return {"passed": passed, "actual": actual, "expected": expected,
            "resolved": addr, "error": None}


def check_output_key_formula(cand: CandidateWorkbook, spec: dict) -> dict:
    """Dereference spec.key via results.json, verify the resolved cell contains a formula."""
    key = spec["key"]
    results = spec.get("_results")
    if results is None:
        return {"passed": False, "actual": None, "expected": "formula",
                "error": "results.json not found next to candidate xlsx"}
    if key not in results:
        return {"passed": False, "actual": None, "expected": "formula",
                "error": f"key '{key}' missing from results.json"}
    addr = results[key]
    try:
        sheet, cell = _parse_cell_addr(addr)
    except ValueError as e:
        return {"passed": False, "actual": addr, "expected": "formula",
                "error": str(e)}
    if sheet not in cand.wb_formulas.sheetnames:
        return {"passed": False, "actual": addr, "expected": "formula",
                "error": f"sheet '{sheet}' not in workbook"}
    raw = cand.wb_formulas[sheet][cell].value
    is_formula = isinstance(raw, str) and raw.startswith("=")
    return {"passed": is_formula, "actual": raw, "expected": "formula",
            "resolved": addr,
            "error": None if is_formula else f"{addr} is {raw!r}, not a formula"}


def check_numeric_anywhere(cand: CandidateWorkbook, spec: dict) -> dict:
    """Does the expected numeric value appear anywhere on the sheet
    (within tolerance)? Useful when both location AND label are flexible.
    Weaker than numeric_by_label — doesn't verify semantic binding.
    """
    sheet = spec["sheet"]
    expected = float(spec["expected"])
    tol_abs = spec.get("tolerance_abs")
    tol_rel = spec.get("tolerance_rel")
    if tol_abs is not None: tol_abs = float(tol_abs)
    if tol_rel is not None: tol_rel = float(tol_rel)
    ws = cand.wb_values[sheet]
    for row in ws.iter_rows():
        for cell in row:
            v = _coerce_number(cell.value)
            if v is not None and _tol_match(v, expected, tol_abs, tol_rel):
                return {"passed": True, "actual": v, "expected": expected,
                        "error": None}
    return {"passed": False, "actual": None, "expected": expected,
            "error": f"value {expected} not found on sheet '{sheet}' within tolerance"}


DISPATCH = {
    "numeric": check_numeric,
    "sheet_exists": check_sheet_exists,
    "label_exists": check_label_exists,
    "formula_driven": check_formula_driven,
    "output_key": check_output_key,
    "output_key_formula": check_output_key_formula,
    "numeric_by_label": check_numeric_by_label,
    "formula_by_label": check_formula_by_label,
    "numeric_anywhere": check_numeric_anywhere,
    "bs_balances": check_bs_balances,
    "hardcode_penalty": check_hardcode_penalty,
    "value_range": check_value_range,
    "range_sum": check_range_sum,
    "llm_judge": check_llm_judge,
}


# ---- public API ------------------------------------------------------------


def grade(
    candidate_xlsx: Path,
    grading_yaml: Path,
    task_dir: Path | None = None,
    recalc: bool | None = None,
) -> dict:
    """Grade a candidate workbook against a rubric.

    task_dir is the task folder, used to resolve relative source_cell paths.
    """
    rubric = yaml.safe_load(Path(grading_yaml).read_text())
    should_recalc = rubric.get("recalc", True) if recalc is None else recalc
    candidate_xlsx = Path(candidate_xlsx)
    if should_recalc:
        recalc_xlsx(candidate_xlsx)

    cand = CandidateWorkbook.load(candidate_xlsx)
    try:
        results_json = _load_results_json(candidate_xlsx)
    except RuntimeError as e:
        results_json = None
        log_err = str(e)
    else:
        log_err = None

    results = []
    weighted = 0.0
    pos_weight = 0.0
    passed = 0
    for spec in rubric["checks"]:
        handler = DISPATCH.get(spec["type"])
        if handler is None:
            results.append({**spec, "passed": False, "error": f"unknown check type: {spec['type']}"})
            continue
        spec_local = dict(spec)
        if task_dir is not None:
            spec_local["_task_dir"] = str(task_dir)
        spec_local["_results"] = results_json
        try:
            r = handler(cand, spec_local)
        except Exception as e:
            r = {"passed": False, "actual": None, "expected": None, "error": f"{type(e).__name__}: {e}"}
        w = float(spec["weight"])
        # Weight semantics:
        #  - positive weight: pass adds w; counts toward pos_weight denominator
        #  - negative weight: treated as penalty — a "fail" (penalty triggers) subtracts |w|
        if w >= 0:
            pos_weight += w
            if r["passed"]:
                weighted += w
                passed += 1
        else:
            if not r["passed"]:
                weighted += w  # w is negative → subtract
        entry = {
            "id": spec["id"], "type": spec["type"], "weight": w,
            "passed": r["passed"], "actual": r.get("actual"),
            "expected": r.get("expected"), "error": r.get("error"),
        }
        if r.get("matched_label") is not None:
            entry["matched_label"] = r["matched_label"]
        if r.get("diagnostics") is not None:
            entry["diagnostics"] = r["diagnostics"]
        results.append(entry)

    total = len(rubric["checks"])
    accuracy = max(0.0, weighted / pos_weight) if pos_weight > 0 else 0.0
    return {
        "accuracy": accuracy,
        "passed": passed,
        "total": total,
        "weighted_score": weighted,
        "total_weight": pos_weight,
        "checks": results,
    }


if __name__ == "__main__":
    import argparse, json
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidate", required=True, type=Path)
    ap.add_argument("--rubric", required=True, type=Path)
    ap.add_argument("--task-dir", type=Path, default=None)
    ap.add_argument("--no-recalc", action="store_true")
    args = ap.parse_args()
    r = grade(args.candidate, args.rubric, task_dir=args.task_dir, recalc=not args.no_recalc)
    print(json.dumps(r, indent=2, default=str))
