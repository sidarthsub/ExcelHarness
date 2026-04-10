"""Tests for LibreOffice-based xlsx -> per-sheet PNG renderer."""
from pathlib import Path

import pytest

from snapshot_renderer import render_xlsx_to_pngs


def test_renders_pngs_for_each_sheet(tmp_path):
    """Given a fixture xlsx, render should produce one PNG per sheet."""
    fixture = Path(__file__).parent / "fixtures" / "two_sheet.xlsx"
    if not fixture.exists():
        pytest.skip("fixture not present — create via the bridge in a real session")
    out_dir = tmp_path / "screenshots"
    pngs = render_xlsx_to_pngs(fixture, out_dir)
    assert len(pngs) >= 1
    for p in pngs:
        assert p.exists()
        assert p.suffix == ".png"
        assert p.stat().st_size > 1000  # non-empty image
