#!/usr/bin/env python3
"""Dump an Excel workbook to flat, grepable TSV files and per-sheet screenshots."""

import json
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from colorsys import hls_to_rgb, rgb_to_hls
from pathlib import Path
from string import ascii_uppercase

import openpyxl
from openpyxl.formula import Tokenizer
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.formula import DataTableFormula
from pdf2image import convert_from_path
from PIL import Image, ImageChops

ROOT = Path(__file__).parent
EVALS = ROOT / "evals"
THEME_ORDER = [
    "lt1", "dk1", "lt2", "dk2",
    "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
    "hlink", "folHlink",
]


def _normalize_hex(value) -> str | None:
    if not value or not isinstance(value, str):
        return None
    text = value.strip().upper()
    if len(text) == 8:
        text = text[-6:]
    if len(text) != 6 or any(ch not in "0123456789ABCDEF" for ch in text):
        return None
    return text


def _apply_tint(hex_rgb: str, tint: float) -> str:
    r = int(hex_rgb[0:2], 16) / 255.0
    g = int(hex_rgb[2:4], 16) / 255.0
    b = int(hex_rgb[4:6], 16) / 255.0
    h, l, s = rgb_to_hls(r, g, b)
    if tint < 0:
        l *= 1.0 + tint
    else:
        l = l * (1.0 - tint) + tint
    r2, g2, b2 = hls_to_rgb(h, l, s)
    return "".join(f"{round(channel * 255):02X}" for channel in (r2, g2, b2))


def _parse_theme_palette(wb) -> dict[int, str]:
    if not wb.loaded_theme:
        return {}
    if isinstance(wb.loaded_theme, bytes):
        theme_xml = wb.loaded_theme
    else:
        theme_xml = wb.loaded_theme.encode("utf-8")
    root = ET.fromstring(theme_xml)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    if scheme is None:
        return {}

    palette = {}
    for idx, key in enumerate(THEME_ORDER):
        el = scheme.find(f"a:{key}", ns)
        if el is None or not list(el):
            continue
        child = list(el)[0]
        raw = child.attrib.get("lastClr") or child.attrib.get("val")
        hex_rgb = _normalize_hex(raw)
        if hex_rgb:
            palette[idx] = hex_rgb
    return palette


def _color_token(color, theme_palette: dict[int, str]) -> str | None:
    if color is None or not getattr(color, "type", None):
        return None

    kind = color.type
    if kind == "rgb":
        hex_rgb = _normalize_hex(color.rgb)
        if not hex_rgb:
            return None
        return f"rgb(#{hex_rgb})"

    if kind == "theme":
        theme_idx = color.theme
        if theme_idx is None:
            return "theme(?)"
        tint = float(color.tint or 0.0)
        resolved = theme_palette.get(theme_idx)
        if resolved:
            resolved = _apply_tint(resolved, tint) if tint else resolved
        if tint:
            if resolved:
                return f"theme({theme_idx},tint={tint:.4f},#{resolved})"
            return f"theme({theme_idx},tint={tint:.4f})"
        if resolved:
            return f"theme({theme_idx},#{resolved})"
        return f"theme({theme_idx})"

    if kind == "indexed":
        index = color.indexed
        if index is None:
            return "indexed(?)"
        if 0 <= index < len(COLOR_INDEX):
            resolved = _normalize_hex(COLOR_INDEX[index])
            if resolved:
                return f"indexed({index},#{resolved})"
        return f"indexed({index})"

    if kind == "auto":
        return "auto"

    return kind


def _cells_ref(row1: int, col1: int, row2: int, col2: int) -> str:
    tl = f"{col_letter(col1)}{row1}"
    br = f"{col_letter(col2)}{row2}"
    return tl if tl == br else f"{tl}:{br}"


def _format_data_table(value: DataTableFormula, anchor_ref: str | None = None) -> str:
    kind = "2D" if getattr(value, "dt2D", None) == "1" else "1D"
    parts = [f"DATA_TABLE[{kind}]"]
    if anchor_ref:
        parts.append(f"anchor:{anchor_ref}")
    if getattr(value, "ref", None):
        parts.append(f"range:{value.ref}")
    if getattr(value, "r1", None):
        parts.append(f"r1:{value.r1}")
    if getattr(value, "r2", None):
        parts.append(f"r2:{value.r2}")
    return " ".join(parts)


def col_letter(n: int) -> str:
    result = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(ascii_uppercase[rem])
    return "".join(reversed(result))


def _parse_sheet_ref(ref: str, current_sheet: str) -> tuple[str, str]:
    if "!" not in ref:
        return current_sheet, ref.replace("$", "")
    sheet_name, cell_ref = ref.rsplit("!", 1)
    sheet_name = sheet_name.strip()
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        sheet_name = sheet_name[1:-1].replace("''", "'")
    return sheet_name, cell_ref.replace("$", "")


def _expand_range_ref(ref: str, current_sheet: str) -> dict:
    sheet_name, cell_ref = _parse_sheet_ref(ref, current_sheet)
    try:
        min_col, min_row, max_col, max_row = range_boundaries(cell_ref)
    except (ValueError, TypeError):
        return {
            "sheet": sheet_name,
            "ref": cell_ref,
            "kind": "unsupported",
            "expanded": [],
        }

    if any(v is None for v in (min_col, min_row, max_col, max_row)):
        return {
            "sheet": sheet_name,
            "ref": cell_ref,
            "kind": "unsupported",
            "expanded": [],
        }

    expanded = [
        f"{col_letter(col)}{row}"
        for row in range(min_row, max_row + 1)
        for col in range(min_col, max_col + 1)
    ]
    kind = "cell" if len(expanded) == 1 else "range"
    return {
        "sheet": sheet_name,
        "ref": cell_ref,
        "kind": kind,
        "expanded": expanded,
    }


def _analyze_formula(formula: str, current_sheet: str) -> dict:
    tokenizer = Tokenizer(formula)
    functions = []
    refs = []
    token_dump = []
    for token in tokenizer.items:
        token_dump.append(
            {
                "type": token.type,
                "subtype": token.subtype,
                "value": token.value,
            }
        )
        if token.type == "FUNC" and token.subtype == "OPEN":
            functions.append(token.value[:-1].upper())
        elif token.type == "OPERAND" and token.subtype == "RANGE":
            refs.append(_expand_range_ref(token.value, current_sheet))

    dependencies = []
    seen = set()
    for ref in refs:
        for cell_ref in ref["expanded"]:
            key = (ref["sheet"], cell_ref)
            if key in seen:
                continue
            seen.add(key)
            dependencies.append(
                {
                    "sheet": ref["sheet"],
                    "cell": cell_ref,
                    "via": ref["ref"],
                }
            )

    return {
        "functions": functions,
        "references": refs,
        "dependencies": dependencies,
        "tokens": token_dump,
    }


def _tarjan_scc(graph: dict[str, set[str]]) -> list[list[str]]:
    index = 0
    stack = []
    indices = {}
    lowlinks = {}
    on_stack = set()
    components = []

    def strongconnect(node: str) -> None:
        nonlocal index
        indices[node] = index
        lowlinks[node] = index
        index += 1
        stack.append(node)
        on_stack.add(node)

        for neighbor in graph.get(node, set()):
            if neighbor not in indices:
                strongconnect(neighbor)
                lowlinks[node] = min(lowlinks[node], lowlinks[neighbor])
            elif neighbor in on_stack:
                lowlinks[node] = min(lowlinks[node], indices[neighbor])

        if lowlinks[node] == indices[node]:
            component = []
            while True:
                other = stack.pop()
                on_stack.remove(other)
                component.append(other)
                if other == node:
                    break
            components.append(sorted(component))

    for node in sorted(graph):
        if node not in indices:
            strongconnect(node)

    return components


def render_screenshots(xlsx_path: Path, sheet_names: list[str], out_dir: Path, source_tag: str = "") -> None:
    """Render per-sheet compressed JPEGs — only for the specified sheets."""
    from snapshot_renderer import render_xlsx_to_pngs
    out_dir.mkdir(parents=True, exist_ok=True)

    # If we're filtering sheets, create a temp xlsx with only those sheets
    # so LibreOffice doesn't render all 33 pages
    all_sheets = openpyxl.load_workbook(xlsx_path, read_only=True).sheetnames
    if set(sheet_names) == set(all_sheets):
        render_path = xlsx_path
    else:
        import shutil
        tmp_xlsx = out_dir / f"_tmp_{xlsx_path.name}"
        shutil.copy2(xlsx_path, tmp_xlsx)
        wb = openpyxl.load_workbook(tmp_xlsx)
        for ws_name in wb.sheetnames:
            if ws_name not in sheet_names:
                del wb[ws_name]
        wb.save(tmp_xlsx)
        wb.close()
        render_path = tmp_xlsx

    outputs = render_xlsx_to_pngs(render_path, out_dir)

    # Map pages to sheet names
    if len(outputs) == len(sheet_names):
        for out_path, name in zip(outputs, sheet_names):
            dest = out_dir / f"{source_tag}{name}.jpg"
            if out_path.exists():
                out_path.rename(dest)
    # Clean up temp files
    for f in out_dir.glob("_tmp_*"):
        f.unlink()
    for f in out_dir.glob("*.pdf"):
        f.unlink()


def dump(xlsx_path: Path, sheets: list[str] | None = None) -> None:
    """Dump workbook to text. If sheets is provided, only dump those sheet names."""
    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        sys.exit(f"File not found: {xlsx_path}")

    # Prefix sheet filenames with source workbook name for easy grepping
    # Skip tag for model.xlsx (post-generation dumps) to keep evaluator lookups simple
    if xlsx_path.stem == "model":
        source_tag = ""
    else:
        source_tag = f"[{xlsx_path.stem}] "

    # 1. Formulas: compact text dump plus structured JSON for machine evaluation
    formulas_dir = EVALS / "formulas"
    formulas_dir.mkdir(parents=True, exist_ok=True)
    # formulas_json_dir disabled — structured JSON was only used by old openpyxl evaluator
    formulas_json_dir = None
    wb = openpyxl.load_workbook(xlsx_path)
    wb_values = openpyxl.load_workbook(xlsx_path, data_only=True)
    theme_palette = _parse_theme_palette(wb)
    sheet_names = [ws.title for ws in wb.worksheets]
    if sheets:
        sheet_set = set(sheets)
        active_sheets = [ws for ws in wb.worksheets if ws.title in sheet_set]
    else:
        active_sheets = list(wb.worksheets)
    all_formula_nodes = set()
    structured_by_sheet = {}
    global_function_counts = defaultdict(int)
    for ws in active_sheets:
        values_ws = wb_values[ws.title]
        max_col = ws.max_column or 1
        actual_max_row = 1
        for r in range(ws.max_row or 1, 0, -1):
            if any(ws.cell(r, c).value is not None for c in range(1, max_col + 1)):
                actual_max_row = r
                break
        max_row = actual_max_row

        lines = []
        structured_formulas = []
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            row_num = row[0].row
            # Collect non-empty cells
            cells = {}  # col_letter -> value
            for cell in row:
                v = cell.value
                if v is not None:
                    cl = col_letter(cell.column)
                    if isinstance(v, DataTableFormula):
                        cells[cl] = _format_data_table(v, cell.coordinate)
                    else:
                        cells[cl] = str(v) if not (isinstance(v, str) and v.startswith("=")) else v

            if not cells:
                continue

            # Separate labels/values from formulas
            labels = {}
            formulas = {}
            for cl, v in cells.items():
                if isinstance(v, str) and v.startswith("="):
                    formulas[cl] = v
                else:
                    labels[cl] = v

            # Try to templatize formulas: replace column letter with {col}
            # Group formulas by their template pattern
            import re as _re
            templates = {}  # template -> list of col letters
            unique_formulas = {}  # col -> formula (not templatizable)
            for cl, f in formulas.items():
                # Replace this column's letter as a relative col ref:
                # I5, I$5, I$19 — but NOT $I5, $I$5 (absolute col refs stay)
                tmpl = _re.sub(rf'(?<!\$)(?<![A-Z]){cl}(?=\$?\d)', '{{col}}', f)
                if tmpl != f:  # successfully templatized
                    templates.setdefault(tmpl, []).append(cl)
                else:
                    unique_formulas[cl] = f

            # Build row output
            parts = []
            if labels:
                label_strs = [f"{cl}={v}" for cl, v in sorted(labels.items())]
                parts.append(" | ".join(label_strs))

            # Group identical unique formulas across consecutive columns
            if unique_formulas:
                groups = []
                sorted_unique = sorted(unique_formulas.items())
                i = 0
                while i < len(sorted_unique):
                    cl, f = sorted_unique[i]
                    run = [cl]
                    while i + 1 < len(sorted_unique) and sorted_unique[i + 1][1] == f:
                        i += 1
                        run.append(sorted_unique[i][0])
                    i += 1
                    if len(run) > 2:
                        parts.append(f"[{run[0]}-{run[-1]}]={f}")
                    else:
                        for c in run:
                            parts.append(f"{c}={f}")

            for tmpl, cols in sorted(templates.items(), key=lambda x: x[1][0]):
                cols_sorted = sorted(cols)
                if len(cols_sorted) == 1:
                    # Single column, write as-is
                    actual = tmpl.replace("{col}", cols_sorted[0])
                    parts.append(f"{cols_sorted[0]}={actual}")
                else:
                    col_range = f"{cols_sorted[0]}-{cols_sorted[-1]}"
                    parts.append(f"[{col_range}]={tmpl}")

            if parts:
                lines.append(f"  {row_num}: {' | '.join(parts)}")

        (formulas_dir / f"{source_tag}{ws.title}.txt").write_text("\n".join(lines))

        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    continue
                node_id = f"{ws.title}!{cell.coordinate}"
                all_formula_nodes.add(node_id)
                analysis = _analyze_formula(cell.value, ws.title)
                for fn_name in analysis["functions"]:
                    global_function_counts[fn_name] += 1
                cached_value = values_ws[cell.coordinate].value
                structured_formulas.append(
                    {
                        "cell": cell.coordinate,
                        "formula": cell.value,
                        "cached_value": cached_value,
                        "functions": analysis["functions"],
                        "references": analysis["references"],
                        "dependencies": analysis["dependencies"],
                        "tokens": analysis["tokens"],
                    }
                )

        structured_by_sheet[ws.title] = structured_formulas

    dependency_graph = {node: set() for node in all_formula_nodes}
    for sheet_name, entries in structured_by_sheet.items():
        for entry in entries:
            node_id = f"{sheet_name}!{entry['cell']}"
            for dep in entry["dependencies"]:
                dep_node = f"{dep['sheet']}!{dep['cell']}"
                if dep_node in all_formula_nodes:
                    dependency_graph[node_id].add(dep_node)

    iterative_components = []
    component_ids = {}
    for component in _tarjan_scc(dependency_graph):
        if len(component) == 1 and component[0] not in dependency_graph.get(component[0], set()):
            continue
        comp_id = f"iter_{len(iterative_components) + 1}"
        iterative_components.append({"id": comp_id, "cells": component})
        for node in component:
            component_ids[node] = comp_id

    # Collect literal cell values referenced by formulas
    literal_cells_by_sheet = defaultdict(dict)
    for sheet_name, entries in structured_by_sheet.items():
        for entry in entries:
            for dep in entry["dependencies"]:
                dep_node = f"{dep['sheet']}!{dep['cell']}"
                if dep_node not in all_formula_nodes:
                    dep_sheet = dep["sheet"]
                    dep_cell = dep["cell"]
                    if dep_cell not in literal_cells_by_sheet[dep_sheet]:
                        try:
                            vs = wb_values[dep_sheet]
                            val = vs[dep_cell].value
                            if val is not None:
                                literal_cells_by_sheet[dep_sheet][dep_cell] = val
                        except (KeyError, AttributeError):
                            pass
            # Also collect literal cells from expanded range references
            for ref in entry.get("references", []):
                ref_sheet = ref.get("sheet", sheet_name)
                for exp_cell in ref.get("expanded", []):
                    exp_node = f"{ref_sheet}!{exp_cell}"
                    if exp_node not in all_formula_nodes and exp_cell not in literal_cells_by_sheet[ref_sheet]:
                        try:
                            vs = wb_values[ref_sheet]
                            val = vs[exp_cell].value
                            if val is not None:
                                literal_cells_by_sheet[ref_sheet][exp_cell] = val
                        except (KeyError, AttributeError):
                            pass

    for sheet_name, entries in structured_by_sheet.items():
        for entry in entries:
            node_id = f"{sheet_name}!{entry['cell']}"
            entry["iterative_component"] = component_ids.get(node_id)
        # Serialize literal cells with consistent types
        literals = {}
        for cell_ref, val in sorted(literal_cells_by_sheet.get(sheet_name, {}).items()):
            if isinstance(val, (int, float)):
                literals[cell_ref] = val
            elif isinstance(val, str):
                literals[cell_ref] = val
            else:
                literals[cell_ref] = str(val)
        payload = {
            "sheet": sheet_name,
            "formula_cells": entries,
            "literal_cells": literals,
        }
        if formulas_json_dir:
            (formulas_json_dir / f"{source_tag}{sheet_name}.json").write_text(
                json.dumps(payload, indent=2, default=str)
            )

    if formulas_json_dir:
        workbook_formula_summary = {
            "workbook": xlsx_path.name,
            "sheet_names": sheet_names,
            "functions_used": dict(sorted(global_function_counts.items())),
            "iterative_components": iterative_components,
        }
        (formulas_json_dir / f"{source_tag}workbook.json").write_text(
            json.dumps(workbook_formula_summary, indent=2, default=str)
        )

    # 2. Styles: formatting metadata per sheet
    styles_dir = EVALS / "styles"
    styles_dir.mkdir(parents=True, exist_ok=True)
    for ws in active_sheets:
        lines = []
        max_col = min(ws.max_column or 1, 30)
        # Find actual last row with data or formatting
        actual_max = 0
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 200), max_col=max_col):
            if any(cell.value is not None or (cell.font and cell.font.bold) or
                   (cell.fill and cell.fill.patternType and cell.fill.patternType != "none") or
                   (cell.border and any(getattr(cell.border, s).style for s in ("top","bottom","left","right") if getattr(cell.border, s)))
                   for cell in row):
                actual_max = row[0].row
        max_row = min(actual_max or 1, 100)

        # Sheet view settings
        show_gridlines = True
        if ws.views and ws.views.sheetView:
            for v in ws.views.sheetView:
                if v.showGridLines is not None:
                    show_gridlines = v.showGridLines
        lines.append(f"## Sheet View")
        lines.append(f"  showGridLines: {show_gridlines}")
        lines.append("")

        # Column widths — rounded. Emit actual width for every column,
        # including sub-1 "spacer" columns (annotated so builders know why
        # they're so narrow — copy the exact number).
        lines.append("## Column Widths")
        for c in range(1, max_col + 1):
            letter = col_letter(c)
            dim = ws.column_dimensions.get(letter)
            if dim and dim.width:
                w = dim.width
                if w < 1.0:
                    lines.append(f"  {letter}: {round(w, 2)}  # spacer — use this exact sub-1 width, do NOT round up")
                else:
                    lines.append(f"  {letter}: {round(w, 1)}")
            else:
                lines.append(f"  {letter}: default")

        # Row heights (non-default only)
        non_default_rows = []
        for r in range(1, max_row + 1):
            dim = ws.row_dimensions.get(r)
            if dim and dim.height and dim.height != 15:  # 15 is Excel default
                non_default_rows.append(f"  {r}: {dim.height}")
        if non_default_rows:
            lines.append("\n## Row Heights (non-default)")
            lines.extend(non_default_rows)

        data_tables = []
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                if isinstance(cell.value, DataTableFormula):
                    data_tables.append(
                        (
                            cell.row,
                            cell.column,
                            (
                                f"  {cell.coordinate}: "
                                f"{_format_data_table(cell.value, cell.coordinate)}"
                            ),
                        )
                    )
        if data_tables:
            lines.append("\n## Data Tables")
            for _, _, line in sorted(data_tables):
                lines.append(line)

        def _cell_base_style_parts(cell) -> tuple[str, ...]:
            """Style token tuple WITHOUT borders."""
            parts = []
            if cell.font:
                has_content = cell.value is not None
                if cell.font.bold:
                    parts.append("bold")
                if cell.font.underline:
                    parts.append(f"underline:{cell.font.underline}")
                # Always surface font family/size on populated cells so theme-font
                # fallback (e.g. Calibri 11 minor) is visible to the evaluator.
                if cell.font.size and (has_content or cell.font.size != 11):
                    parts.append(f"size:{cell.font.size}")
                if cell.font.name and (has_content or cell.font.name != "Calibri"):
                    parts.append(f"font:{cell.font.name}")
                if getattr(cell.font, "scheme", None) and (
                    has_content or getattr(cell.font, "scheme", None) != "minor"
                ):
                    parts.append(f"scheme:{cell.font.scheme}")
                font_color = _color_token(cell.font.color, theme_palette)
                if font_color and (has_content or font_color != "theme(1,#000000)"):
                    parts.append(f"color:{font_color}")
            if cell.fill and cell.fill.patternType and cell.fill.patternType != "none":
                fill_color = _color_token(cell.fill.fgColor, theme_palette)
                if fill_color:
                    parts.append(f"fill:{fill_color}")
                else:
                    parts.append(f"fill:{cell.fill.patternType}")
            if cell.number_format and cell.number_format != "General":
                parts.append(f"fmt:{cell.number_format}")
            if cell.alignment:
                if cell.alignment.horizontal:
                    parts.append(f"align:{cell.alignment.horizontal}")
                if cell.alignment.indent and cell.alignment.indent > 0:
                    parts.append(f"indent:{cell.alignment.indent}")
            return tuple(parts)

        def _cell_borders(cell) -> dict:
            """Return {side: {style, color}} for non-empty borders."""
            result = {}
            if cell.border:
                for side in ("top", "bottom", "left", "right"):
                    b = getattr(cell.border, side)
                    if b and b.style:
                        result[side] = {
                            "style": b.style,
                            "color": _color_token(b.color, theme_palette),
                        }
            return result

        def _cell_alignment_parts(cell) -> tuple[str, ...]:
            """Alignment token tuple for non-default alignment behavior."""
            parts = []
            if cell.alignment:
                if cell.alignment.horizontal:
                    parts.append(f"h:{cell.alignment.horizontal}")
                if cell.alignment.vertical and cell.alignment.vertical != "bottom":
                    parts.append(f"v:{cell.alignment.vertical}")
                if cell.alignment.indent and cell.alignment.indent > 0:
                    parts.append(f"indent:{cell.alignment.indent}")
                if cell.alignment.wrap_text:
                    parts.append("wrap")
                if cell.alignment.shrink_to_fit:
                    parts.append("shrink")
                if cell.alignment.text_rotation:
                    parts.append(f"rotate:{cell.alignment.text_rotation}")
            return tuple(parts)

        def _merge_rectangles(pos_to_value: dict[tuple[int, int], str]) -> list[tuple[int, int, str]]:
            """Merge same-valued positions into simple rectangles."""
            value_to_cells = defaultdict(set)
            for pos, value in pos_to_value.items():
                value_to_cells[value].add(pos)

            range_lines = []
            for value in sorted(value_to_cells.keys()):
                cell_set = set(value_to_cells[value])
                for (r, c) in sorted(value_to_cells[value]):
                    if (r, c) not in cell_set:
                        continue
                    max_c = c
                    while (r, max_c + 1) in cell_set:
                        max_c += 1
                    max_r = r
                    while all((max_r + 1, cc) in cell_set for cc in range(c, max_c + 1)):
                        max_r += 1
                    for rr in range(r, max_r + 1):
                        for cc in range(c, max_c + 1):
                            cell_set.discard((rr, cc))
                    ref = _cells_ref(r, c, max_r, max_c)
                    range_lines.append((r, c, f"  {ref}: {value}"))
            range_lines.sort()
            return range_lines

        # Pass 1: collect base styles (no borders) and borders separately
        style_grid = {}  # (row, col) -> base style token tuple
        font_grid = {}  # (row, col) -> font signature for populated cells
        alignment_grid = {}  # (row, col) -> alignment signature
        border_grid = {}  # (row, col) -> {side: {style, color}}
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                base = _cell_base_style_parts(cell)
                if base:
                    style_grid[(cell.row, cell.column)] = base
                if cell.value is not None and cell.font:
                    font_parts = []
                    if cell.font.name:
                        font_parts.append(f"font:{cell.font.name}")
                    if cell.font.size:
                        font_parts.append(f"size:{cell.font.size}")
                    if getattr(cell.font, "scheme", None):
                        font_parts.append(f"scheme:{cell.font.scheme}")
                    if cell.font.bold:
                        font_parts.append("bold")
                    if cell.font.underline:
                        font_parts.append(f"underline:{cell.font.underline}")
                    if font_parts:
                        font_grid[(cell.row, cell.column)] = " ".join(font_parts)
                alignment = _cell_alignment_parts(cell)
                if alignment:
                    alignment_grid[(cell.row, cell.column)] = " ".join(alignment)
                borders = _cell_borders(cell)
                if borders:
                    border_grid[(cell.row, cell.column)] = borders

        # Pass 2: build style table (deduplicate)
        # Simplify color tokens: strip wrapper, keep just hex
        import re as _re
        def _simplify_token(t: str) -> str:
            # color:indexed(0,#000000) → color:#000000
            t = _re.sub(r'color:indexed\(\d+,#([0-9A-Fa-f]{6})\)', r'color:#\1', t)
            t = _re.sub(r'color:indexed\(\d+\)', 'color:#000000', t)
            t = _re.sub(r'color:rgb\(#([0-9A-Fa-f]{6})\)', r'color:#\1', t)
            t = _re.sub(r'color:theme\(\d+,#([0-9A-Fa-f]{6})\)', r'color:#\1', t)
            t = _re.sub(r'color:theme\(\d+,tint=[^,]+,#([0-9A-Fa-f]{6})\)', r'color:#\1', t)
            t = _re.sub(r'fill:rgb\(#([0-9A-Fa-f]{6})\)', r'fill:#\1', t)
            t = _re.sub(r'fill:theme\(\d+,#([0-9A-Fa-f]{6})\)', r'fill:#\1', t)
            return t

        unique_styles = sorted(set(style_grid.values()))

        # Detect base font (most common font+size combo) and factor it out
        from collections import Counter as _Counter
        font_tokens = _Counter()
        for style in unique_styles:
            font_parts = tuple(t for t in style if t.startswith("font:") or t.startswith("size:"))
            if font_parts:
                font_tokens[font_parts] += 1
        base_font = font_tokens.most_common(1)[0][0] if font_tokens else ()
        base_font_set = set(base_font)

        # Simplify and deduplicate tokens, stripping base font from each style
        simplified_styles = []
        for style in unique_styles:
            simplified = tuple(_simplify_token(t) for t in style if t not in base_font_set)
            simplified_styles.append(simplified)

        style_ids = {s: f"S{i+1}" for i, s in enumerate(unique_styles)}
        simplified_style_ids = dict(zip(unique_styles, simplified_styles))

        # Collect tokens from simplified styles
        all_tokens = sorted({t for s in simplified_styles for t in s})
        token_ids = {t: f"T{i+1}" for i, t in enumerate(all_tokens)}

        if base_font:
            lines.append(f"\n## Base Font (applied to all cells, not repeated in styles)")
            lines.append(f"  {' | '.join(base_font)}")

        if token_ids:
            lines.append("\n## Style Tokens")
            for token, tid in sorted(token_ids.items(), key=lambda item: int(item[1][1:])):
                lines.append(f"  {tid}: {token}")

        lines.append("\n## Style Table")
        for style, sid in sorted(style_ids.items(), key=lambda x: x[1]):
            simplified = simplified_style_ids[style]
            if simplified:
                tokens = " ".join(token_ids[t] for t in simplified)
                lines.append(f"  {sid}: {tokens}")
            else:
                lines.append(f"  {sid}: (base only)")

        # Font Map and Alignment Map dropped — redundant with Cell Styles

        # Pass 3: map cells to style IDs and merge into 2D rectangles
        id_grid = {pos: style_ids[s] for pos, s in style_grid.items()}

        lines.append("\n## Cell Styles")
        lines.extend(line for _, _, line in _merge_rectangles(id_grid))

        # Pass 4: borders — exact edge spans with deduped specs plus semantic overlays
        edge_cells = defaultdict(set)  # (side, style, color) -> set of (row, col)
        for (r, c_), borders in border_grid.items():
            for side, meta in borders.items():
                edge_cells[(side, meta["style"], meta["color"] or "-")].add((r, c_))

        edge_spans = []

        def _append_edge_span(side: str, style: str, color: str, row1: int, col1: int, row2: int, col2: int) -> None:
            edge_spans.append((side, style, color, row1, col1, row2, col2))

        for (side, style, color), cells in sorted(edge_cells.items()):
            if side in ("top", "bottom"):
                rows = defaultdict(list)
                for r, c in cells:
                    rows[r].append(c)
                for r in sorted(rows):
                    cols = sorted(rows[r])
                    start = prev = cols[0]
                    for c in cols[1:]:
                        if c == prev + 1:
                            prev = c
                            continue
                        _append_edge_span(side, style, color, r, start, r, prev)
                        start = prev = c
                    _append_edge_span(side, style, color, r, start, r, prev)
            else:
                cols = defaultdict(list)
                for r, c in cells:
                    cols[c].append(r)
                for c in sorted(cols):
                    rows = sorted(cols[c])
                    start = prev = rows[0]
                    for r in rows[1:]:
                        if r == prev + 1:
                            prev = r
                            continue
                        _append_edge_span(side, style, color, start, c, prev, c)
                        start = prev = r
                    _append_edge_span(side, style, color, start, c, prev, c)

        edge_spans.sort(key=lambda item: (item[3], item[4], item[0], item[1], item[2], item[5], item[6]))

        # Weight mapping for Office.js
        weight_map = {"hair": "Hairline", "thin": "Thin", "medium": "Medium", "thick": "Thick"}

        def _color_to_hex(color_token: str) -> str:
            """Convert color token to hex color for bridge commands."""
            if color_token == "-" or color_token == "auto":
                return "#000000"
            if color_token.startswith("rgb(#"):
                return color_token[4:-1]  # extract #XXXXXX
            if color_token.startswith("indexed("):
                # Extract hex if present, otherwise default to black
                if ",#" in color_token:
                    return color_token.split(",#")[1].rstrip(")")
                return "#000000"
            if color_token.startswith("theme("):
                if ",#" in color_token:
                    return color_token.split(",#")[1].rstrip(")")
                return "#000000"
            return "#000000"

        # Compact border format: group sides per range, one line each
        # Format: RANGE SIDES WEIGHT [COLOR]  (color omitted if black)
        from collections import defaultdict as _dd
        border_groups = _dd(list)  # (range, weight, color) -> [sides]
        for side, style, color, row1, col1, row2, col2 in edge_spans:
            ref = _cells_ref(row1, col1, row2, col2)
            hex_color = _color_to_hex(color)
            weight = weight_map.get(style, "thin")
            border_groups[(ref, weight.lower(), hex_color)].append(side)

        if border_groups:
            lines.append("\n## Borders")
            lines.append("# Format: RANGE SIDES WEIGHT [COLOR if not black]")
            for (ref, weight, color), sides in sorted(border_groups.items()):
                sides_str = "+".join(sorted(set(sides)))
                color_str = f" {color}" if color != "#000000" else ""
                lines.append(f"  {ref} {sides_str} {weight}{color_str}")

        (styles_dir / f"{source_tag}{ws.title}.txt").write_text("\n".join(lines))

    # 5. Screenshots: real LibreOffice renders via PDF → PNG
    screenshots_dir = EVALS / "screenshots"
    active_names = [ws.title for ws in active_sheets]
    render_screenshots(xlsx_path, active_names, screenshots_dir, source_tag)

    print(f"Dump complete: {EVALS} (tagged: {source_tag.strip()})")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("xlsx", nargs="?", default=str(ROOT / "models" / "model.xlsx"))
    parser.add_argument("--sheets", help="Comma-separated sheet names to dump (default: all)")
    args = parser.parse_args()
    sheet_filter = [s.strip() for s in args.sheets.split(",")] if args.sheets else None
    dump(Path(args.xlsx), sheets=sheet_filter)
