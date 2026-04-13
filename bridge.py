"""Synchronous HTTP client that Builder scripts import to drive live Excel.

This is the interface the Builder agent writes code against. Keep it thin —
one function per bridge command, no convention encoding, no helpers.
"""
from __future__ import annotations

import requests
import urllib3


class BridgeError(Exception):
    pass


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
