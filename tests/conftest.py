"""Pytest configuration and fixtures."""

import sys
from pathlib import Path

# Add project root to sys.path so tests can import project modules
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import and re-export fixtures and helpers for all tests
from .helpers import tmp_xlsx

__all__ = ["tmp_xlsx"]
