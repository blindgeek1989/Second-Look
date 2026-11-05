"""
Step 4 tests — slide rendering via the adapter interface.

Tests that run without LibreOffice
-----------------------------------
- Session creation, uniqueness, and cleanup (including cleanup on exception)
- File validation: missing file, oversized file, non-ZIP file

Tests that need LibreOffice (skipped automatically if not installed)
---------------------------------------------------------------------
- Each fixture renders to the right number of PNG files
- Output files are valid PNGs (correct magic bytes)
- Output files are named slide_001.png, slide_002.png, …
- Output dimensions are reasonable for a 10 × 7.5 inch slide
- Session deletes all rendered images on exit

Install LibreOffice and set LIBREOFFICE_PATH in .env to run the full suite.
On Windows the default path is:
  C:\\Program Files\\LibreOffice\\program\\soffice.exe
"""

from __future__ import annotations

import shutil
import struct
from pathlib import Path

import pytest

from app.config import LIBREOFFICE_PATH, MAX_UPLOAD_MB
from app.rendering import (
    CorruptFileError,
    FileTooLargeError,
    LibreOfficeRenderer,
    RenderError,
    render_session,
)

FIXTURES = Path(__file__).parent / "fixtures"


# ── LibreOffice availability check ───────────────────────────────────────────

def _lo_available() -> bool:
    p = Path(LIBREOFFICE_PATH)
    return p.is_file() or shutil.which(LIBREOFFICE_PATH) is not None


requires_lo = pytest.mark.skipif(
    not _lo_available(),
    reason=(
        f"LibreOffice not found at '{LIBREOFFICE_PATH}'. "
        "Install LibreOffice and set LIBREOFFICE_PATH in .env to run rendering tests."
    ),
)

ALL_FIXTURES = [
    "clean.pptx",
    "overlapping_objects.pptx",
    "object_off_slide.pptx",
    "text_overflow.pptx",
    "small_text.pptx",
    "low_contrast.pptx",
    "missing_alt_text.pptx",
    "no_title.pptx",
    "reading_order.pptx",
    "crowded.pptx",
]


# ── Session lifecycle (no LibreOffice needed) ─────────────────────────────────

def test_session_creates_directory(tmp_path):
    with render_session(tmp_path) as session_dir:
        assert session_dir.is_dir()


def test_session_cleans_up_on_normal_exit(tmp_path):
    with render_session(tmp_path) as session_dir:
        captured = session_dir
    assert not captured.exists()


def test_session_cleans_up_on_exception(tmp_path):
    captured = None
    try:
        with render_session(tmp_path) as session_dir:
            captured = session_dir
            # Create a file inside to confirm everything gets deleted
            (session_dir / "test.txt").write_text("hello")
            raise RuntimeError("simulated failure")
    except RuntimeError:
        pass
    assert captured is not None
    assert not captured.exists()


def test_session_each_call_gets_unique_directory(tmp_path):
    names = []
    for _ in range(5):
        with render_session(tmp_path) as d:
            names.append(d.name)
    assert len(set(names)) == 5, "Expected every session to have a unique directory name"


# ── File validation (no LibreOffice needed) ───────────────────────────────────

def test_missing_file_raises_render_error(tmp_path):
    renderer = LibreOfficeRenderer()
    with pytest.raises(RenderError):
        renderer.render(tmp_path / "does_not_exist.pptx", tmp_path)


def test_oversized_file_raises_file_too_large(tmp_path):
    oversized = tmp_path / "big.pptx"
    # Write exactly 1 byte over the limit
    oversized.write_bytes(b"\x00" * (MAX_UPLOAD_MB * 1024 * 1024 + 1))
    renderer = LibreOfficeRenderer()
    with pytest.raises(FileTooLargeError):
        renderer.render(oversized, tmp_path)


def test_non_zip_file_raises_corrupt_error(tmp_path):
    garbage = tmp_path / "corrupt.pptx"
    garbage.write_bytes(b"THIS IS NOT A ZIP FILE AND NEVER WILL BE")
    renderer = LibreOfficeRenderer()
    with pytest.raises(CorruptFileError):
        renderer.render(garbage, tmp_path)


def test_file_at_exact_limit_is_accepted(tmp_path):
    """A file exactly at the limit should not raise FileTooLargeError.
    It will raise CorruptFileError (not a ZIP) — but not the size error."""
    at_limit = tmp_path / "exact.pptx"
    at_limit.write_bytes(b"\x00" * (MAX_UPLOAD_MB * 1024 * 1024))
    renderer = LibreOfficeRenderer()
    with pytest.raises(RenderError) as exc_info:
        renderer.render(at_limit, tmp_path)
    assert not isinstance(exc_info.value, FileTooLargeError)


# ── LibreOffice rendering (skipped if not installed) ─────────────────────────

@requires_lo
@pytest.mark.parametrize("fixture", ALL_FIXTURES)
def test_fixture_renders_one_png_per_slide(fixture, tmp_path):
    renderer = LibreOfficeRenderer()
    with render_session(tmp_path) as session_dir:
        images = renderer.render(FIXTURES / fixture, session_dir)
    # All test fixtures are single-slide decks
    assert len(images) == 1, (
        f"{fixture}: expected 1 image, got {len(images)}"
    )
    assert images[0].suffix == ".png"


@requires_lo
def test_rendered_png_has_correct_magic_bytes(tmp_path):
    renderer = LibreOfficeRenderer()
    with render_session(tmp_path) as session_dir:
        images = renderer.render(FIXTURES / "clean.pptx", session_dir)
        # PNG magic: 8 bytes starting with \x89PNG
        header = images[0].read_bytes()[:8]
    assert header == b"\x89PNG\r\n\x1a\n"


@requires_lo
def test_rendered_images_named_sequentially(tmp_path):
    renderer = LibreOfficeRenderer()
    with render_session(tmp_path) as session_dir:
        images = renderer.render(FIXTURES / "clean.pptx", session_dir)
    assert images[0].name == "slide_001.png"


@requires_lo
def test_rendered_image_dimensions_are_reasonable(tmp_path):
    """
    A 10 × 7.5 inch slide at 2× scale (144 DPI) should be roughly
    1440 × 1080 px. We allow ±10% for LibreOffice rounding.
    """
    renderer = LibreOfficeRenderer()
    with render_session(tmp_path) as session_dir:
        images = renderer.render(FIXTURES / "clean.pptx", session_dir)
        # Read PNG dimensions from the IHDR chunk (bytes 16–24)
        data = images[0].read_bytes()

    width = struct.unpack(">I", data[16:20])[0]
    height = struct.unpack(">I", data[20:24])[0]

    assert 1200 <= width <= 1700, f"Unexpected width: {width}px"
    assert 900 <= height <= 1300, f"Unexpected height: {height}px"


@requires_lo
def test_session_deletes_rendered_images_on_exit(tmp_path):
    renderer = LibreOfficeRenderer()
    captured: list[Path] = []

    with render_session(tmp_path) as session_dir:
        images = renderer.render(FIXTURES / "clean.pptx", session_dir)
        captured.extend(images)

    for p in captured:
        assert not p.exists(), f"Expected '{p.name}' to be deleted after session exit"


@requires_lo
def test_session_deletes_images_even_on_exception(tmp_path):
    renderer = LibreOfficeRenderer()
    captured: list[Path] = []

    try:
        with render_session(tmp_path) as session_dir:
            images = renderer.render(FIXTURES / "clean.pptx", session_dir)
            captured.extend(images)
            raise RuntimeError("simulated downstream failure")
    except RuntimeError:
        pass

    for p in captured:
        assert not p.exists(), f"Expected '{p.name}' to be deleted despite exception"
