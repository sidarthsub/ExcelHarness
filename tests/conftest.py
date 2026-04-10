"""Shared fixtures for the v3 harness test suite."""
import pytest
from pathlib import Path

HARNESS_ROOT = Path(__file__).parent.parent


@pytest.fixture
def harness_root():
    return HARNESS_ROOT
