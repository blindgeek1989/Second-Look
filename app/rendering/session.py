"""
Session management for Second Look rendering.

Each upload gets a unique temporary directory. render_session() guarantees
that the directory and everything in it is deleted when the block exits,
even if an exception is raised. This satisfies Product Rule 10:
"Delete uploaded files and generated images after a session ends."
"""

from __future__ import annotations

import shutil
import uuid
from contextlib import contextmanager
from pathlib import Path


@contextmanager
def render_session(base_dir: Path | None = None):
    """
    Create a temporary session directory and delete it on exit.

    Parameters
    ----------
    base_dir:
        Parent directory for session folders. Defaults to ``sessions/``
        in the current working directory.

    Yields
    ------
    Path
        The session directory (guaranteed empty and writable on entry).

    Example
    -------
    ::

        with render_session() as session_dir:
            images = renderer.render(pptx_path, session_dir)
            # use images here
        # session_dir and all images are gone here
    """
    if base_dir is None:
        base_dir = Path("sessions")

    session_dir = base_dir / uuid.uuid4().hex
    session_dir.mkdir(parents=True, exist_ok=True)

    try:
        yield session_dir
    finally:
        shutil.rmtree(session_dir, ignore_errors=True)
