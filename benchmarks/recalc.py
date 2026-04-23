"""Recalculate an xlsx's cached values via headless LibreOffice.

openpyxl writes formulas but does not compute them. The grader compares
numeric cell values against gold values, which requires real computed
values. We shell out to soffice --headless --calc --convert-to xlsx to
force a recalc and cache the results.

Falls back to returning the input path unchanged if soffice isn't
available — the grader will then use whatever cached values openpyxl
preserves, which may be None for freshly-written formulas.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path


def soffice_path() -> str | None:
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if Path(mac).exists():
        return mac
    return None


def recalc_xlsx(src: Path, dest: Path | None = None, timeout: float = 60.0) -> Path:
    """Recalculate src into an xlsx with populated cached values.

    If dest is None, overwrites src in place (via a tmp file).
    Returns the path to the recalculated file (dest or src).
    """
    src = Path(src)
    if dest is None:
        dest = src
    dest = Path(dest)

    exe = soffice_path()
    if exe is None:
        return src  # no-op fallback

    outdir = dest.parent / ".recalc_tmp"
    outdir.mkdir(parents=True, exist_ok=True)
    try:
        result = subprocess.run(
            [
                exe,
                "--headless",
                "--calc",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(outdir),
                str(src),
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if result.returncode != 0:
            raise RuntimeError(f"soffice failed: {result.stderr[:400]}")
        produced = outdir / (src.stem + ".xlsx")
        if not produced.exists():
            raise RuntimeError(f"soffice produced no file in {outdir}")
        shutil.move(str(produced), str(dest))
        return dest
    finally:
        if outdir.exists():
            shutil.rmtree(outdir, ignore_errors=True)
