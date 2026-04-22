"""Synchronous HTTP client that Builder scripts import to drive live Excel.

This is the interface the Builder agent writes code against. Keep it thin —
one function per bridge command, no convention encoding, no helpers.
"""
from __future__ import annotations

import requests
import urllib3


class BridgeError(Exception):
    pass


def _copy_cell_format(cell) -> dict:
    """Extract a format dict from an openpyxl cell matching the bridge's formatRange schema.

    Skips default values to keep the payload minimal. Used by copy_sheet_from_input.
    """
    fmt: dict = {}
    font: dict = {}
    if cell.font.name and cell.font.name != "Calibri":
        font["name"] = cell.font.name
    if cell.font.size and cell.font.size != 11.0:
        font["size"] = float(cell.font.size)
    if cell.font.bold:
        font["bold"] = True
    if cell.font.italic:
        font["italic"] = True
    color = cell.font.color
    if color is not None and getattr(color, "rgb", None):
        rgb = color.rgb
        if isinstance(rgb, str) and len(rgb) >= 6:
            hex_rgb = rgb[-6:].upper()
            if hex_rgb != "000000":
                font["color"] = f"#{hex_rgb}"
    if font:
        fmt["font"] = font

    if cell.fill.patternType == "solid":
        fg = cell.fill.fgColor
        if fg is not None and getattr(fg, "rgb", None):
            rgb = fg.rgb
            if isinstance(rgb, str) and len(rgb) >= 6:
                hex_rgb = rgb[-6:].upper()
                if hex_rgb != "FFFFFF":
                    fmt["fill"] = {"color": f"#{hex_rgb}"}

    align = cell.alignment.horizontal
    if align in ("center", "right", "left", "centerContinuous"):
        mapping = {
            "center": "Center",
            "right": "Right",
            "left": "Left",
            "centerContinuous": "CenterAcrossSelection",
        }
        fmt["horizontalAlignment"] = mapping[align]

    return fmt


def _freeze(obj):
    """Recursively convert a dict/list into a hashable tuple form for dict keys."""
    if isinstance(obj, dict):
        return tuple(sorted((k, _freeze(v)) for k, v in obj.items()))
    if isinstance(obj, list):
        return tuple(_freeze(v) for v in obj)
    return obj


def _thaw(obj):
    """Inverse of _freeze — reconstruct dict/list form."""
    if isinstance(obj, tuple) and obj and isinstance(obj[0], tuple) and len(obj[0]) == 2 and isinstance(obj[0][0], str):
        return {k: _thaw(v) for k, v in obj}
    if isinstance(obj, tuple):
        return [_thaw(v) for v in obj]
    return obj


class Bridge:
    def __init__(self, base_url: str = "https://localhost:3000", verify_tls: bool = True, timeout: float = 60.0):
        self.base_url = base_url.rstrip("/")
        self.verify = verify_tls
        self.timeout = timeout
        if not verify_tls:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def _call(self, command: str, params: dict | None = None) -> dict:
        resp = requests.post(
            f"{self.base_url}/api/command",
            json={"command": command, "params": params or {}},
            verify=self.verify,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        body = resp.json()
        if "error" in body and not body.get("ok"):
            raise BridgeError(f"{command} failed: {body['error']}")
        return body

    # --- Sheet management ---
    def create_sheet(self, name: str) -> dict:
        return self._call("createSheet", {"name": name})

    def delete_sheet(self, name: str) -> dict:
        return self._call("deleteSheet", {"name": name})

    def get_sheet_names(self) -> dict:
        return self._call("getSheetNames", {})

    # --- Values and formulas ---
    def write_values(self, sheet: str, address: str, values: list) -> dict:
        return self._call("writeValues", {"sheet": sheet, "address": address, "values": values})

    def write_formulas(self, sheet: str, address: str, formulas: list) -> dict:
        return self._call("writeFormulas", {"sheet": sheet, "address": address, "formulas": formulas})

    def read_values(self, sheet: str, address: str) -> dict:
        return self._call("readValues", {"sheet": sheet, "address": address})

    def read_format(self, sheet: str, address: str) -> dict:
        return self._call("readFormat", {"sheet": sheet, "address": address})

    # --- Formatting ---
    def format_range(self, sheet: str, address: str, format: dict) -> dict:
        return self._call("formatRange", {"sheet": sheet, "address": address, "format": format})

    def set_number_format(self, sheet: str, address: str, format: str) -> dict:
        return self._call("setNumberFormat", {"sheet": sheet, "address": address, "format": format})

    def set_column_widths(self, sheet: str, columns: dict) -> dict:
        return self._call("setColumnWidths", {"sheet": sheet, "columns": columns})

    def set_row_heights(self, sheet: str, rows: dict) -> dict:
        return self._call("setRowHeights", {"sheet": sheet, "rows": rows})

    def auto_fit_columns(self, sheet: str, address: str | None = None) -> dict:
        return self._call("autoFitColumns", {"sheet": sheet, "address": address})

    def auto_fit_rows(self, sheet: str, address: str | None = None) -> dict:
        return self._call("autoFitRows", {"sheet": sheet, "address": address})

    def merge_cells(self, sheet: str, address: str) -> dict:
        return self._call("mergeCells", {"sheet": sheet, "address": address})

    def freeze_rows(self, sheet: str, count: int) -> dict:
        return self._call("freezeRows", {"sheet": sheet, "count": count})

    def set_show_gridlines(self, sheet: str, show: bool) -> dict:
        return self._call("setShowGridLines", {"sheet": sheet, "show": show})

    def clear_range(self, sheet: str, address: str) -> dict:
        return self._call("clearRange", {"sheet": sheet, "address": address})

    # --- Charts ---
    def create_chart(self, sheet: str, **kwargs) -> dict:
        return self._call("createChart", {"sheet": sheet, **kwargs})

    # --- Workbook settings ---
    def set_iterative_calculation(self, enabled: bool, max_iteration: int = 100, max_change: float = 0.001) -> dict:
        return self._call("setIterativeCalculation", {
            "enabled": enabled,
            "maxIteration": max_iteration,
            "maxChange": max_change,
        })

    def create_table(self, sheet: str, header_address: str, name: str, rows: list, style: str | None = None) -> dict:
        params = {"sheet": sheet, "headerAddress": header_address, "name": name, "rows": rows}
        if style:
            params["style"] = style
        return self._call("createTable", params)

    # --- Batch ---
    def batch(self, commands: list) -> dict:
        return self._call("batch", {"commands": commands})

    def dump_sheet(self, sheet: str) -> dict:
        return self._call("dumpSheet", {"sheet": sheet})

    def copy_sheet_from_input(
        self,
        xlsx_path: str,
        source_sheet: str,
        target_sheet: str | None = None,
    ) -> dict:
        """Copy a sheet from an input xlsx into the live workbook via openpyxl.

        Replaces `create_sheet` + dozens of write/format calls for "this sheet
        is an exact copy of the input" cases. Handles values, formulas,
        column widths, row heights, gridlines, and per-cell formatting
        (font, fill, alignment, number format). Idempotent — deletes the
        target sheet if it already exists.

        Args:
            xlsx_path: absolute path to the source .xlsx.
            source_sheet: sheet name within the xlsx to copy.
            target_sheet: name for the new live sheet. Defaults to source_sheet.
        """
        import openpyxl
        from collections import defaultdict
        from openpyxl.utils import get_column_letter
        from pathlib import Path

        path = Path(xlsx_path).expanduser()
        if not path.exists():
            raise BridgeError(f"copy_sheet_from_input: file not found: {path}")

        wb = openpyxl.load_workbook(path, data_only=False)
        if source_sheet not in wb.sheetnames:
            raise BridgeError(
                f"copy_sheet_from_input: sheet '{source_sheet}' not in {path.name} "
                f"(available: {wb.sheetnames})"
            )
        ws = wb[source_sheet]
        target = target_sheet or source_sheet

        self.create_sheet(target)

        if ws.sheet_view.showGridLines is False:
            self.set_show_gridlines(target, False)

        widths = {c: d.width for c, d in ws.column_dimensions.items() if d.width}
        if widths:
            self.set_column_widths(target, widths)

        heights = {str(r): d.height for r, d in ws.row_dimensions.items() if d.height}
        if heights:
            self.set_row_heights(target, heights)

        max_row = ws.max_row or 0
        max_col = ws.max_column or 0
        if not (max_row and max_col):
            wb.close()
            return {"ok": True, "target": target, "note": "empty source sheet"}

        values = [[None] * max_col for _ in range(max_row)]
        formulas = [[None] * max_col for _ in range(max_row)]
        has_values = False
        has_formulas = False
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                r, c = cell.row - 1, cell.column - 1
                if isinstance(v, str) and v.startswith("="):
                    formulas[r][c] = v
                    has_formulas = True
                else:
                    values[r][c] = v
                    has_values = True

        full_range = f"A1:{get_column_letter(max_col)}{max_row}"
        if has_values:
            self.write_values(target, full_range, values)
        if has_formulas:
            self.write_formulas(target, full_range, formulas)

        # Group cells by identical format signature → one format_range call per group.
        format_groups: dict[tuple, list[str]] = defaultdict(list)
        numfmt_groups: dict[str, list[str]] = defaultdict(list)
        for row in ws.iter_rows():
            for cell in row:
                fmt = _copy_cell_format(cell)
                if fmt:
                    format_groups[_freeze(fmt)].append(cell.coordinate)
                nf = cell.number_format
                if nf and nf != "General":
                    numfmt_groups[nf].append(cell.coordinate)

        batch_cmds = []
        for fmt_key, coords in format_groups.items():
            fmt = _thaw(fmt_key)
            batch_cmds.append({
                "command": "formatRange",
                "params": {"sheet": target, "address": ",".join(coords), "format": fmt},
            })
        for nf, coords in numfmt_groups.items():
            batch_cmds.append({
                "command": "setNumberFormat",
                "params": {"sheet": target, "address": ",".join(coords), "format": nf},
            })

        if batch_cmds:
            self.batch(batch_cmds)

        wb.close()
        return {
            "ok": True,
            "target": target,
            "format_groups": len(format_groups),
            "numfmt_groups": len(numfmt_groups),
        }

    def protect_workbook(self, password: str | None = None) -> dict:
        return self._call("protectWorkbook", {"password": password})

    def unprotect_workbook(self, password: str | None = None) -> dict:
        return self._call("unprotectWorkbook", {"password": password})

    def save_snapshot(self) -> dict:
        return self._call("saveSnapshot", {})

    def set_status(self, text: str) -> dict:
        return self._call("setStatus", {"text": text})

    def checkpoint(self, description: str) -> dict:
        """Block until the harness resolves this checkpoint (after evaluator + commit)."""
        resp = requests.post(
            f"{self.base_url}/api/checkpoint",
            json={"description": description},
            verify=self.verify,
            timeout=600,  # long poll — checkpoint may take a while
        )
        resp.raise_for_status()
        return resp.json()

    def emit(self, text: str) -> dict:
        """Send an agent-initiated chat message to the taskpane."""
        resp = requests.post(
            f"{self.base_url}/api/emit",
            json={"text": text},
            verify=self.verify,
            timeout=self.timeout,
        )
        resp.raise_for_status()
        return resp.json()
