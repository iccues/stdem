"""
Pytest configuration and shared fixtures
"""

import pytest
from pathlib import Path


@pytest.fixture(scope="session")
def test_excel_dir():
    """Path to test Excel files directory"""
    return Path("tests/excel")


@pytest.fixture(scope="session")
def test_json_dir():
    """Path to test JSON files directory"""
    return Path("tests/json")
