"""
LibreOffice slide rendering adapter for Second Look.

Pipeline: PPTX → PDF (LibreOffice headless) → PNG per page (PyMuPDF).

This is the pilot adapter. The rendering interface (SlideRenderer) is
defined in base.py so this implementation can be replaced without
touching the analysis module, AI review, or web interface.
"""

from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path

import fitz  # PyMuPDF

from app.config import (
    CONVERSION_TIMEOUT_SECONDS,
    LIBREOFFICE_PATH,
    MAX_UPLOAD_MB,
)
from app.rendering.base import (
    CorruptFileError,
    ConversionTimeoutError,
    FileTooLargeError,
    RenderError,
    SlideRenderer,
)

# 2× zoom on a 72 DPI PDF base → 144 DPI output images.
# A standard 10 × 7.5 inch slide becomes 1 440 × 1 080 px — enough detail
# for the Anthropic vision model while keeping file sizes manageable.
_IMAGE_SCALE = 2.0
_MAX_BYTES = MAX_UPLOAD_MB * 1024 * 1024


class LibreOfficeRenderer(SlideRenderer):
    """
    Renders PPTX slides to PNG via LibreOffice and PyMuPDF.

    Requires LibreOffice to be installed and the LIBREOFFICE_PATH
    environment variable to point to the soffice executable.
    """

    def render(self, pptx_path: Path, output_dir: Path) -> list[Path]:
        pptx_path = Path(pptx_path)
        output_dir = Path(output_dir)

        self._validate(pptx_path)
        pdf_path = self._to_pdf(pptx_path, output_dir)
        images = self._to_images(pdf_path, output_dir)
        pdf_path.unlink(missing_ok=True)  # intermediate PDF no longer needed
        return images

    # ── Validation ────────────────────────────────────────────────────────

    def _validate(self, path: Path) -> None:
        if not path.exists():
            raise RenderError(f"File not found: {path}")

        if path.stat().st_size > _MAX_BYTES:
            mb = path.stat().st_size / (1024 * 1024)
            raise FileTooLargeError(
                f"'{path.name}' is {mb:.1f} MB; "
                f"the upload limit is {MAX_UPLOAD_MB} MB."
            )

        # PPTX is a ZIP archive. Catching a non-ZIP early gives a clearer
        # error message than waiting for LibreOffice to fail.
        if not zipfile.is_zipfile(str(path)):
            raise CorruptFileError(
                f"'{path.name}' does not appear to be a valid PPTX file "
                f"(failed ZIP structure check)."
            )

    # ── PPTX → PDF ────────────────────────────────────────────────────────

    def _to_pdf(self, pptx_path: Path, output_dir: Path) -> Path:
        cmd = [
            LIBREOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(pptx_path),
        ]

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=CONVERSION_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired:
            raise ConversionTimeoutError(
                f"LibreOffice did not finish converting '{pptx_path.name}' "
                f"within {CONVERSION_TIMEOUT_SECONDS} seconds."
            )
        except FileNotFoundError:
            raise RenderError(
                f"LibreOffice executable not found at '{LIBREOFFICE_PATH}'. "
                f"Install LibreOffice and set LIBREOFFICE_PATH in your .env file."
            )

        if result.returncode != 0:
            raise RenderError(
                f"LibreOffice exited with code {result.returncode} "
                f"while converting '{pptx_path.name}'. "
                f"stderr: {result.stderr.strip()!r}"
            )

        pdf_path = output_dir / (pptx_path.stem + ".pdf")
        if not pdf_path.exists():
            raise RenderError(
                f"LibreOffice reported success but '{pdf_path.name}' "
                f"was not created in the output directory."
            )

        return pdf_path

    # ── PDF → PNG ─────────────────────────────────────────────────────────

    def _to_images(self, pdf_path: Path, output_dir: Path) -> list[Path]:
        try:
            doc = fitz.open(str(pdf_path))
        except Exception as exc:
            raise RenderError(
                f"Could not open '{pdf_path.name}' for image conversion: {exc}"
            ) from exc

        images: list[Path] = []
        matrix = fitz.Matrix(_IMAGE_SCALE, _IMAGE_SCALE)

        try:
            for i, page in enumerate(doc):
                img_path = output_dir / f"slide_{i + 1:03d}.png"
                pix = page.get_pixmap(matrix=matrix)
                pix.save(str(img_path))
                images.append(img_path)
        finally:
            doc.close()

        if not images:
            raise RenderError(
                f"'{pdf_path.name}' produced no pages after conversion."
            )

        return images
