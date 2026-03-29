#!/usr/bin/env python3
"""Dump an Excel workbook to flat, grepable TSV files and per-sheet screenshots."""

import subprocess
import sys
import tempfile
from pathlib import Path
from string import ascii_uppercase

import openpyxl
from pdf2image import convert_from_path
from PIL import Image

ROOT = Path(__file__).parent
EVALS = ROOT / "evals"


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

        # Cell formatting — style table + 2D ranges (borders reported separately)
        from collections import defaultdict

        def _cell_base_style(cell) -> str:
            """Style string WITHOUT borders."""
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
                if cell.font.color and cell.font.color.type == "rgb" and cell.font.color.rgb:
                    rgb = str(cell.font.color.rgb)
                    if len(rgb) >= 6 and rgb not in ("00000000", "FF000000", "0"):
                        parts.append(f"color:#{rgb[-6:]}")
            if cell.fill and cell.fill.patternType and cell.fill.patternType != "none":
                if cell.fill.fgColor and cell.fill.fgColor.rgb:
                    rgb = str(cell.fill.fgColor.rgb)
                    if rgb not in ("00000000", "0"):
                        parts.append(f"fill:#{rgb[-6:]}")
            if cell.number_format and cell.number_format != "General":
                parts.append(f"fmt:{cell.number_format}")
            if cell.alignment:
                if cell.alignment.horizontal:
                    parts.append(f"align:{cell.alignment.horizontal}")
                if cell.alignment.indent and cell.alignment.indent > 0:
                    parts.append(f"indent:{cell.alignment.indent}")
            return " | ".join(parts)

        def _cell_borders(cell) -> dict:
            """Return {side: style} for non-empty borders."""
            result = {}
            if cell.border:
                for side in ("top", "bottom", "left", "right"):
                    b = getattr(cell.border, side)
                    if b and b.style:
                        result[side] = b.style
            return result

        # Pass 1: collect base styles (no borders) and borders separately
        style_grid = {}  # (row, col) -> base style string
        border_grid = {}  # (row, col) -> {side: style}
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                base = _cell_base_style(cell)
                if base:
                    style_grid[(cell.row, cell.column)] = base
                borders = _cell_borders(cell)
                if borders:
                    border_grid[(cell.row, cell.column)] = borders

        # Pass 2: build style table (deduplicate)
        unique_styles = sorted(set(style_grid.values()))
        style_ids = {s: f"S{i+1}" for i, s in enumerate(unique_styles)}

        lines.append("\n## Style Table")
        for style, sid in sorted(style_ids.items(), key=lambda x: x[1]):
            lines.append(f"  {sid}: {style}")

        # Pass 3: map cells to style IDs and merge into 2D rectangles
        id_grid = {pos: style_ids[s] for pos, s in style_grid.items()}

        lines.append("\n## Cell Styles")
        id_to_cells = defaultdict(set)
        for pos, sid in id_grid.items():
            id_to_cells[sid].add(pos)

        range_lines = []
        for sid in sorted(id_to_cells.keys(), key=lambda x: int(x[1:])):
            cell_set = set(id_to_cells[sid])
            for (r, c) in sorted(id_to_cells[sid]):
                if (r, c) not in cell_set:
                    continue
                # Expand right
                max_c = c
                while (r, max_c + 1) in cell_set:
                    max_c += 1
                # Expand down (full column span must match)
                max_r = r
                while all((max_r + 1, cc) in cell_set for cc in range(c, max_c + 1)):
                    max_r += 1
                # Mark used
                for rr in range(r, max_r + 1):
                    for cc in range(c, max_c + 1):
                        cell_set.discard((rr, cc))
                tl = f"{col_letter(c)}{r}"
                br = f"{col_letter(max_c)}{max_r}"
                ref = tl if tl == br else f"{tl}:{br}"
                range_lines.append((r, c, f"  {ref}: {sid}"))

        range_lines.sort()
        lines.extend(line for _, _, line in range_lines)

        # Pass 4: borders — merge into row/column spans
        lines.append("\n## Borders")
        # Group by (side, style) and find contiguous runs
        edge_cells = defaultdict(set)  # (side, style) -> set of (row, col)
        for (r, c_), borders in border_grid.items():
            for side, bstyle in borders.items():
                edge_cells[(side, bstyle)].add((r, c_))

        border_lines = []
        for (side, bstyle), cells in sorted(edge_cells.items()):
            cell_set = set(cells)
            for (r, c) in sorted(cells):
                if (r, c) not in cell_set:
                    continue
                # For top/bottom borders, expand horizontally
                # For left/right borders, expand vertically
                if side in ("top", "bottom"):
                    max_c = c
                    while (r, max_c + 1) in cell_set:
                        max_c += 1
                    # Also try expanding down
                    max_r = r
                    while all((max_r + 1, cc) in cell_set for cc in range(c, max_c + 1)):
                        max_r += 1
                    for rr in range(r, max_r + 1):
                        for cc in range(c, max_c + 1):
                            cell_set.discard((rr, cc))
                    tl = f"{col_letter(c)}{r}"
                    br = f"{col_letter(max_c)}{max_r}"
                else:
                    max_r = r
                    while (max_r + 1, c) in cell_set:
                        max_r += 1
                    # Also try expanding right
                    max_c = c
                    while all((rr, max_c + 1) in cell_set for rr in range(r, max_r + 1)):
                        max_c += 1
                    for rr in range(r, max_r + 1):
                        for cc in range(c, max_c + 1):
                            cell_set.discard((rr, cc))
                    tl = f"{col_letter(c)}{r}"
                    br = f"{col_letter(max_c)}{max_r}"
                ref = tl if tl == br else f"{tl}:{br}"
                border_lines.append((r, c, f"  {ref}: {side}:{bstyle}"))

        border_lines.sort()
        lines.extend(line for _, _, line in border_lines)

        (styles_dir / f"{source_tag}{ws.title}.txt").write_text("\n".join(lines))

    # 5. Screenshots: real LibreOffice renders via PDF → PNG
    screenshots_dir = EVALS / "screenshots"
    render_screenshots(xlsx_path, sheet_names, screenshots_dir, source_tag)

    print(f"Dump complete: {EVALS} (tagged: {source_tag.strip()})")


if __name__ == "__main__":
    model = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "models" / "model.xlsx"
    dump(model)
