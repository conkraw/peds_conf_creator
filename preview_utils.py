from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


def office_binary() -> str | None:
    for candidate in ["libreoffice", "soffice", "/usr/bin/libreoffice", "/usr/bin/soffice"]:
        found = shutil.which(candidate) if not candidate.startswith("/") else candidate
        if found and Path(found).exists():
            return found
    return None


def _run_libreoffice_convert(input_path: Path, outdir: Path, convert_to: str, timeout: int = 120) -> subprocess.CompletedProcess | None:
    office = office_binary()
    if not office:
        return None
    env = dict(os.environ)
    env["HOME"] = str(outdir)
    env["UserInstallation"] = f"file://{outdir / 'lo-profile'}"
    cmd = [
        office,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--nolockcheck",
        "--convert-to",
        convert_to,
        "--outdir",
        str(outdir),
        str(input_path),
    ]
    return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout, env=env)


def _render_pdf_first_page_to_png(pdf_path: Path) -> bytes | None:
    if fitz is None or not pdf_path.exists():
        return None
    try:
        doc = fitz.open(str(pdf_path))
        if doc.page_count < 1:
            doc.close()
            return None
        page = doc.load_page(0)
        matrix = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        data = pix.tobytes("png")
        doc.close()
        return data
    except Exception:
        return None


def render_pptx_first_slide_to_png(pptx_bytes: bytes) -> bytes | None:
    if not pptx_bytes:
        return None
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp = Path(tmpdir)
            input_path = tmp / "uploaded_slide.pptx"
            input_path.write_bytes(pptx_bytes)

            pdf_result = _run_libreoffice_convert(input_path, tmp, "pdf")
            if pdf_result is not None and pdf_result.returncode == 0:
                pdf_candidates = sorted(tmp.glob("*.pdf"))
                if pdf_candidates:
                    rendered = _render_pdf_first_page_to_png(pdf_candidates[0])
                    if rendered:
                        return rendered

            png_result = _run_libreoffice_convert(input_path, tmp, "png")
            if png_result is not None and png_result.returncode == 0:
                png_candidates = sorted(tmp.glob("*.png"))
                if png_candidates:
                    return png_candidates[0].read_bytes()
    except Exception:
        return None
    return None
