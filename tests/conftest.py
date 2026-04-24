"""Pytest configuration and fixtures."""

import sys
from pathlib import Path

# Add project root to sys.path so tests can import core.*, web.*, cli.*
sys.path.insert(0, str(Path(__file__).parent.parent))

from .helpers import tmp_xlsx

__all__ = ["tmp_xlsx"]
