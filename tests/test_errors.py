"""
Tests for error handling and validation
"""

import unittest
from src import stdem
from src.stdem.TableException import (
    TableFileNotFoundError,
    InvalidFileFormatError,
    MissingHeaderMarkerError,
    MissingDataMarkerError,
    InvalidTypeNameError,
    InvalidHeaderFormatError,
)
from tests.test_base import BaseTestCase, TemporaryExcelFile


class TestFileErrors(BaseTestCase):
    """Test file-related errors"""

    def test_empty_filename(self):
        """Test that empty filename raises ValueError"""
        with self.assertRaises(ValueError) as context:
            stdem.ExcelParser.getData("")
        self.assertIn("cannot be empty", str(context.exception))

    def test_nonexistent_file(self):
        """Test that nonexistent file raises FileNotFoundError"""
        with self.assertRaises(TableFileNotFoundError) as context:
            stdem.ExcelParser.getData("tests/excel/nonexistent.xlsx")
        self.assertIn("not found", str(context.exception))

    def test_invalid_file_format(self):
        """Test that non-xlsx file raises InvalidFileFormatError"""
        with self.assertRaises(InvalidFileFormatError):
            stdem.ExcelParser.getData("tests/test_errors.py")


class TestHeaderErrors(BaseTestCase):
    """Test header-related errors"""

    def test_missing_head_marker(self):
        """Test that file without #head marker raises error"""

        def setup(wb, ws):
            ws["A1"] = "invalid"
            ws["B1"] = "name:string"

        with TemporaryExcelFile("test_no_head.xlsx", setup) as test_file:
            with self.assertRaises(MissingHeaderMarkerError) as context:
                stdem.ExcelParser.getData(test_file)
            self.assertIn("#head", str(context.exception))

    def test_invalid_type_name(self):
        """Test that invalid type name raises error"""

        def setup(wb, ws):
            ws["A1"] = "#head"
            ws["B1"] = "name:invalid_type"

        with TemporaryExcelFile("test_invalid_type.xlsx", setup) as test_file:
            with self.assertRaises(InvalidTypeNameError) as context:
                stdem.ExcelParser.getData(test_file)
            self.assertIn("Invalid type", str(context.exception))

    def test_invalid_header_format(self):
        """Test that header without colon raises error"""

        def setup(wb, ws):
            ws["A1"] = "#head"
            ws["B1"] = "namestring"  # Missing colon

        with TemporaryExcelFile("test_invalid_format.xlsx", setup) as test_file:
            with self.assertRaises(InvalidHeaderFormatError) as context:
                stdem.ExcelParser.getData(test_file)
            self.assertIn("format", str(context.exception).lower())


class TestDataErrors(BaseTestCase):
    """Test data-related errors"""

    def test_missing_data_marker(self):
        """Test that file without #data marker raises error"""

        def setup(wb, ws):
            ws["A1"] = "#head"
            ws["B1"] = "name:string"
            # No #data row

        with TemporaryExcelFile("test_no_data.xlsx", setup) as test_file:
            with self.assertRaises(MissingDataMarkerError) as context:
                stdem.ExcelParser.getData(test_file)
            self.assertIn("#data", str(context.exception))


if __name__ == "__main__":
    unittest.main()
