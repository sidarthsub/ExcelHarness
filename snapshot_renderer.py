"""Render an .xlsx into per-sheet PNGs via LibreOffice + pdftoppm.

Used by the Evaluator to get visual feedback on the live workbook's state
at checkpoint time. Office.js has no worksheet image export, so we go
via a saved snapshot + LibreOffice.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


def render_xlsx_to_pngs(xlsx_path: Path, out_dir: Path) -> list[Path]:
    """Convert xlsx -> PDF -> per-sheet PNGs.

    Returns the list of produced PNG paths, in page order.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: xlsx -> PDF via LibreOffice headless
    pdf_path = out_dir / f"{xlsx_path.stem}.pdf"
    result = subprocess.run(
        [
            "soffice", "--headless",
            "--convert-to", "pdf",
            "--outdir", str(out_dir),
            str(xlsx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0 or not pdf_path.exists():
        raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

    # Step 2: PDF -> PNGs via pdftoppm
    png_prefix = out_dir / f"{xlsx_path.stem}_page"
    result = subprocess.run(
        ["pdftoppm", "-png", "-r", "150", str(pdf_path), str(png_prefix)],
        capture_output=True,
        text=True,
        timeout=60,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed: {result.stderr}")

    pngs = sorted(out_dir.glob(f"{xlsx_path.stem}_page-*.png"))
    return pngs


def ensure_tools_available() -> None:
    """Raise if soffice or pdftoppm is not on PATH."""
    for tool in ("soffice", "pdftoppm"):
        if shutil.which(tool) is None:
            raise RuntimeError(
                f"{tool} not found on PATH. "
                "Install LibreOffice (brew install --cask libreoffice) "
                "and poppler (brew install poppler)."
            )
