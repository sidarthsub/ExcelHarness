#!/usr/bin/env python3
"""Dump an Excel workbook to flat, grepable TSV files and per-sheet screenshots."""

import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from colorsys import hls_to_rgb, rgb_to_hls
from pathlib import Path
from string import ascii_uppercase

import openpyxl
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.worksheet.formula import DataTableFormula
from pdf2image import convert_from_path
from PIL import Image

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


def render_screenshots(xlsx_path: Path, sheet_names: list[str], out_dir: Path, source_tag: str = "") -> None:
    """Export per-sheet PNGs via LibreOffice PDF export + pdf2image."""
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(tmpdir), str(xlsx_path)],
            check=True, capture_output=True,
        )
        # Find the generated PDF
        pdfs = list(tmpdir.glob("*.pdf"))
        if not pdfs:
            print("  Warning: PDF export failed, skipping screenshots", file=sys.stderr)
            return
        images = convert_from_path(str(pdfs[0]), dpi=150)

        if len(images) == len(sheet_names):
            for name, img in zip(sheet_names, images):
                img.save(str(out_dir / f"{source_tag}{name}.png"))
        elif len(images) > len(sheet_names):
            # Wide sheets span multiple pages — stitch vertically per sheet
            pages_per_sheet = len(images) // len(sheet_names)
            remainder = len(images) % len(sheet_names)
            idx = 0
            for i, name in enumerate(sheet_names):
                count = pages_per_sheet + (1 if i < remainder else 0)
                sheet_imgs = images[idx:idx + count]
                idx += count
                if len(sheet_imgs) == 1:
                    sheet_imgs[0].save(str(out_dir / f"{name}.png"))
                else:
                    total_h = sum(im.height for im in sheet_imgs)
                    max_w = max(im.width for im in sheet_imgs)
                    stitched = Image.new("RGB", (max_w, total_h), (255, 255, 255))
                    y = 0
                    for im in sheet_imgs:
                        stitched.paste(im, (0, y))
                        y += im.height
                    stitched.save(str(out_dir / f"{name}.png"))
        else:
            for i, img in enumerate(images):
                name = sheet_names[i] if i < len(sheet_names) else f"page_{i}"
                img.save(str(out_dir / f"{source_tag}{name}.png"))


def dump(xlsx_path: Path) -> None:
    xlsx_path = Path(xlsx_path).resolve()
    if not xlsx_path.exists():
        sys.exit(f"File not found: {xlsx_path}")

    # Prefix sheet filenames with source workbook name for easy grepping
    # Skip tag for model.xlsx (post-generation dumps) to keep evaluator lookups simple
    if xlsx_path.stem == "model":
        source_tag = ""
    else:
        source_tag = f"[{xlsx_path.stem}] "

    # 1. Formulas: compact format with row templates for repeated formulas
    formulas_dir = EVALS / "formulas"
    formulas_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(xlsx_path)
    theme_palette = _parse_theme_palette(wb)
    sheet_names = [ws.title for ws in wb.worksheets]
    for ws in wb.worksheets:
        max_col = ws.max_column or 1
        actual_max_row = 1
        for r in range(ws.max_row or 1, 0, -1):
            if any(ws.cell(r, c).value is not None for c in range(1, max_col + 1)):
                actual_max_row = r
                break
        max_row = actual_max_row

        lines = []
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

            if unique_formulas:
                for cl, f in sorted(unique_formulas.items()):
                    parts.append(f"{cl}={f}")

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

    # 2. Styles: formatting metadata per sheet
    styles_dir = EVALS / "styles"
    styles_dir.mkdir(parents=True, exist_ok=True)
    for ws in wb.worksheets:
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

        # Column widths
        lines.append("## Column Widths")
        for c in range(1, max_col + 1):
            letter = col_letter(c)
            dim = ws.column_dimensions.get(letter)
            w = dim.width if dim and dim.width else "default"
            lines.append(f"  {letter}: {w}")

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
                if cell.font.bold:
                    parts.append("bold")
                if cell.font.underline:
                    parts.append(f"underline:{cell.font.underline}")
                if cell.font.size and cell.font.size != 11:
                    parts.append(f"size:{cell.font.size}")
                if cell.font.name and cell.font.name != "Calibri":
                    parts.append(f"font:{cell.font.name}")
                font_color = _color_token(cell.font.color, theme_palette)
                if font_color:
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
        alignment_grid = {}  # (row, col) -> alignment signature
        border_grid = {}  # (row, col) -> {side: {style, color}}
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                base = _cell_base_style_parts(cell)
                if base:
                    style_grid[(cell.row, cell.column)] = base
                alignment = _cell_alignment_parts(cell)
                if alignment:
                    alignment_grid[(cell.row, cell.column)] = " ".join(alignment)
                borders = _cell_borders(cell)
                if borders:
                    border_grid[(cell.row, cell.column)] = borders

        # Pass 2: build style table (deduplicate)
        unique_styles = sorted(set(style_grid.values()))
        style_ids = {s: f"S{i+1}" for i, s in enumerate(unique_styles)}
        style_tokens = sorted({token for style in unique_styles for token in style})
        token_ids = {token: f"T{i+1}" for i, token in enumerate(style_tokens)}

        if token_ids:
            lines.append("\n## Style Tokens")
            for token, tid in sorted(token_ids.items(), key=lambda item: int(item[1][1:])):
                lines.append(f"  {tid}: {token}")

        lines.append("\n## Style Table")
        for style, sid in sorted(style_ids.items(), key=lambda x: x[1]):
            tokens = " ".join(token_ids[token] for token in style)
            lines.append(f"  {sid}: {tokens}")

        if alignment_grid:
            lines.append("\n## Alignment Map")
            for _, _, line in _merge_rectangles(alignment_grid):
                lines.append(line)

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

        border_colors = sorted({color for _, _, color, *_ in edge_spans if color != "-"})
        color_ids = {color: f"BC{i + 1}" for i, color in enumerate(border_colors)}

        spec_keys = sorted({(side, style, color) for side, style, color, *_ in edge_spans})
        spec_ids = {spec: f"BS{i + 1}" for i, spec in enumerate(spec_keys)}

        if color_ids:
            lines.append("\n## Border Colors")
            for color, color_id in sorted(color_ids.items(), key=lambda item: item[1]):
                lines.append(f"  {color_id}: {color}")

        if spec_ids:
            lines.append("\n## Border Specs")
            for (side, style, color), spec_id in sorted(spec_ids.items(), key=lambda item: item[1]):
                color_part = f" {color_ids[color]}" if color in color_ids else ""
                lines.append(f"  {spec_id}: {side} {style}{color_part}")

        lines.append("\n## Borders Raw")
        refs_by_spec = defaultdict(list)
        top_by_sig = defaultdict(list)
        bottom_by_sig = defaultdict(list)
        left_lookup = {}
        right_lookup = {}

        for side, style, color, row1, col1, row2, col2 in edge_spans:
            spec_id = spec_ids[(side, style, color)]
            ref = _cells_ref(row1, col1, row2, col2)
            refs_by_spec[spec_id].append((row1, col1, ref))
            if side == "top":
                top_by_sig[(style, color, col1, col2)].append((row1, spec_id, ref))
            elif side == "bottom":
                bottom_by_sig[(style, color, col1, col2)].append((row1, spec_id, ref))
            elif side == "left":
                left_lookup[(style, color, col1, row1, row2)] = (spec_id, ref)
            elif side == "right":
                right_lookup[(style, color, col1, row1, row2)] = (spec_id, ref)

        for spec_id in sorted(refs_by_spec.keys(), key=lambda value: int(value[2:])):
            refs = " | ".join(ref for _, _, ref in sorted(refs_by_spec[spec_id]))
            lines.append(f"  {spec_id}: {refs}")

        lines.append("\n## Structural Regions")
        structural_lines = []
        seen_outlines = set()
        outline_specs = set()

        for (style, color, col1, col2), tops in sorted(top_by_sig.items()):
            bottoms = sorted(bottom_by_sig.get((style, color, col1, col2), []))
            if not bottoms:
                continue
            for top_row, top_spec, top_ref in sorted(tops):
                for bottom_row, bottom_spec, bottom_ref in bottoms:
                    if bottom_row <= top_row:
                        continue
                    left_meta = left_lookup.get((style, color, col1, top_row, bottom_row))
                    right_meta = right_lookup.get((style, color, col2, top_row, bottom_row))
                    if not left_meta or not right_meta:
                        continue
                    outline_key = (top_row, col1, bottom_row, col2, style, color)
                    if outline_key in seen_outlines:
                        continue
                    seen_outlines.add(outline_key)
                    ref = _cells_ref(top_row, col1, bottom_row, col2)
                    left_spec, left_ref = left_meta
                    right_spec, right_ref = right_meta
                    outline_specs.update({top_spec, bottom_spec, left_spec, right_spec})
                    structural_lines.append(
                        (
                            top_row,
                            col1,
                            (
                                f"  region kind:outline range:{ref} "
                                f"top:{top_spec}@{top_ref} "
                                f"bottom:{bottom_spec}@{bottom_ref} "
                                f"left:{left_spec}@{left_ref} "
                                f"right:{right_spec}@{right_ref}"
                            ),
                        )
                    )

        spec_by_id = {spec_id: spec for spec, spec_id in spec_ids.items()}
        for spec_id in sorted(refs_by_spec.keys(), key=lambda value: int(value[2:])):
            if spec_id in outline_specs:
                continue
            side, style, color = spec_by_id[spec_id]
            if side in ("top", "bottom"):
                kind = "hline"
            else:
                kind = "vline"
            refs = " | ".join(ref for _, _, ref in sorted(refs_by_spec[spec_id]))
            color_part = f" color:{color_ids[color]}" if color in color_ids else ""
            structural_lines.append(
                (
                    9999,
                    int(spec_id[2:]),
                    f"  region kind:{kind} edge:{side} spec:{spec_id} style:{style}{color_part} ranges:{refs}",
                )
            )

        structural_lines.sort()
        lines.extend(line for _, _, line in structural_lines)

        (styles_dir / f"{source_tag}{ws.title}.txt").write_text("\n".join(lines))

    # 5. Screenshots: real LibreOffice renders via PDF → PNG
    screenshots_dir = EVALS / "screenshots"
    render_screenshots(xlsx_path, sheet_names, screenshots_dir, source_tag)

    print(f"Dump complete: {EVALS} (tagged: {source_tag.strip()})")


if __name__ == "__main__":
    model = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "models" / "model.xlsx"
    dump(model)
