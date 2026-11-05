"""
Slide rendering interface for Second Look.

Any adapter must implement SlideRenderer.render().
The rest of the application depends only on this contract — not on LibreOffice,
Microsoft Graph, or any other specific renderer.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class RenderError(Exception):
    """Base class for all rendering failures."""


class FileTooLargeError(RenderError):
    """File exceeds the configured upload size limit."""


class CorruptFileError(RenderError):
    """File is not a valid or supported format."""


class ConversionTimeoutError(RenderError):
    """Rendering process exceeded the configured timeout."""


class SlideRenderer(ABC):
    """
    Abstract interface for slide rendering adapters.

    Contract
    --------
    Accepts a PPTX file path and a session output directory.
    Returns a list of PNG image paths in slide order (slide 1 first).

    Raises
    ------
    FileTooLargeError       — file exceeds MAX_UPLOAD_MB
    CorruptFileError        — file is not a valid PPTX
    ConversionTimeoutError  — rendering exceeded the timeout
    RenderError             — any other rendering failure
    """

    @abstractmethod
    def render(self, pptx_path: Path, output_dir: Path) -> list[Path]:
        """Render all slides to PNG images and return their paths."""
