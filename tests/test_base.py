"""
Base test utilities and fixtures
"""
import unittest
import os
import openpyxl
from pathlib import Path


class BaseTestCase(unittest.TestCase):
    """Base test case with common utilities"""

    @classmethod
    def setUpClass(cls):
        """Set up test fixtures"""
        cls.test_excel_dir = Path("tests/excel")
        cls.test_json_dir = Path("tests/json")

    def create_test_excel(self, filename: str, setup_func=None):
        """
        Create a test Excel file

        Args:
            filename: Name of the test file
            setup_func: Function to set up the workbook content
                       Takes (workbook, worksheet) as arguments

        Returns:
            Path to the created file
        """
        wb = openpyxl.Workbook()
        ws = wb.active

        if setup_func:
            setup_func(wb, ws)

        test_file = self.test_excel_dir / filename
        wb.save(str(test_file))
        return test_file

    def cleanup_test_file(self, filepath):
        """Remove a test file if it exists"""
        if filepath and os.path.exists(filepath):
            os.remove(filepath)


class TemporaryExcelFile:
    """Context manager for temporary Excel files"""

    def __init__(self, filename: str, setup_func=None):
        self.filename = Path("tests/excel") / filename
        self.setup_func = setup_func

    def __enter__(self):
        wb = openpyxl.Workbook()
        ws = wb.active

        if self.setup_func:
            self.setup_func(wb, ws)

        wb.save(str(self.filename))
        return str(self.filename)

    def __exit__(self, exc_type, exc_val, exc_tb):
        if os.path.exists(self.filename):
            os.remove(self.filename)
