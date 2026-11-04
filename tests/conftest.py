"""
pytest configuration — auto-generate fixtures before tests run if needed.
"""

import subprocess
import sys
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"
MANIFEST = FIXTURES_DIR / "manifest.json"


def pytest_configure(config):
    """Generate fixture files if they do not yet exist."""
    if not MANIFEST.exists():
        script = Path(__file__).parent / "create_fixtures.py"
        print(f"\nFixtures not found — generating from {script.name} …")
        subprocess.run([sys.executable, str(script)], check=True)
