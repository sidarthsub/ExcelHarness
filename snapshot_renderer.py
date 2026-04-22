"""Render an .xlsx into per-sheet compressed JPEGs via LibreOffice + pdf2image.

Unified renderer — used by both dump.py (input screenshots) and the evaluator
(checkpoint screenshots). Auto-crops whitespace and compresses to ~30-60KB per sheet.
"""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from PIL import Image, ImageChops
from pdf2image import convert_from_path
import numpy as np


MAX_WIDTH = 800


def _save_compressed(img: Image.Image, out_path: Path) -> None:
    """Autocrop whitespace, resize, save as JPEG."""
    arr = np.array(img)
    row_mins = arr.min(axis=(1, 2))
    col_mins = arr.min(axis=(0, 2))
    content_rows = np.where(row_mins < 240)[0]
    content_cols = np.where(col_mins < 240)[0]
    if len(content_rows) > 0 and len(content_cols) > 0:
        pad = 10
        top = max(0, content_rows[0] - pad)
        bottom = min(arr.shape[0], content_rows[-1] + pad)
        left = max(0, content_cols[0] - pad)
        right = min(arr.shape[1], content_cols[-1] + pad)
        img = img.crop((left, top, right, bottom))
    if img.width > MAX_WIDTH:
        ratio = MAX_WIDTH / img.width
        img = img.resize((MAX_WIDTH, int(img.height * ratio)), Image.LANCZOS)
    img.save(str(out_path), "JPEG", quality=55)


def render_xlsx_to_pngs(xlsx_path: Path, out_dir: Path) -> list[Path]:
    """Convert xlsx -> PDF -> per-sheet compressed JPEGs.

    Returns the list of produced JPEG paths, in page order.
    Despite the function name, outputs are now JPEGs for size efficiency.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", str(tmpdir), str(xlsx_path)],
            capture_output=True, text=True, timeout=120,
        )
        pdfs = list(tmpdir.glob("*.pdf"))
        if not pdfs:
            raise RuntimeError(f"LibreOffice conversion failed: {result.stderr}")

        # Also copy PDF to out_dir for reference
        pdf_dest = out_dir / f"{xlsx_path.stem}.pdf"
        shutil.copy2(pdfs[0], pdf_dest)

        images = convert_from_path(str(pdfs[0]), dpi=100)

    outputs = []
    for i, img in enumerate(images):
        out_path = out_dir / f"{xlsx_path.stem}_page-{i+1:02d}.jpg"
        _save_compressed(img, out_path)
        outputs.append(out_path)

    return outputs


def ensure_tools_available() -> None:
    """Raise if soffice is not on PATH."""
    if shutil.which("soffice") is None:
        raise RuntimeError(
            "soffice not found on PATH. "
            "Install LibreOffice (brew install --cask libreoffice)."
        )
