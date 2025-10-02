"""
Tests for basic Excel parsing functionality
"""
import unittest
import json
from src import stdem
from tests.test_base import BaseTestCase


class TestBasicParsing(BaseTestCase):
    """Test basic parsing of Excel files"""

    def test_example_excel(self):
        """Test parsing example.xlsx and compare with expected JSON"""
        result = stdem.ExcelParser.getData("tests/excel/example.xlsx")

        with open("tests/json/example.json", "r", encoding="utf-8") as f:
            expected = json.load(f)

        self.assertEqual(result, expected)

    def test_unit_data_excel(self):
        """Test parsing UnitData.xlsx and compare with expected JSON"""
        result = stdem.ExcelParser.getData("tests/excel/UnitData.xlsx")

        with open("tests/json/UnitData.json", "r", encoding="utf-8") as f:
            expected = json.load(f)

        self.assertEqual(result, expected)

    def test_skill_table_excel(self):
        """Test parsing SkillTable.xlsx and compare with expected JSON"""
        result = stdem.ExcelParser.getData("tests/excel/SkillTable.xlsx")

        with open("tests/json/SkillTable.json", "r", encoding="utf-8") as f:
            expected = json.load(f)

        self.assertEqual(result, expected)

    def test_effect_table_excel(self):
        """Test parsing EffectTable.xlsx and compare with expected JSON"""
        result = stdem.ExcelParser.getData("tests/excel/EffectTable.xlsx")

        with open("tests/json/EffectTable.json", "r", encoding="utf-8") as f:
            expected = json.load(f)

        self.assertEqual(result, expected)


class TestJSONFormatting(BaseTestCase):
    """Test JSON output formatting"""

    def test_getJson_returns_formatted_json(self):
        """Test that getJson returns formatted JSON with indentation"""
        json_str = stdem.ExcelParser.getJson("tests/excel/example.xlsx")

        # Should be valid JSON
        parsed = json.loads(json_str)
        self.assertIsInstance(parsed, dict)

        # Should have indentation
        self.assertIn("\n", json_str)

        # Compare with expected
        with open("tests/json/example.json", "r", encoding="utf-8") as f:
            expected = json.load(f)
        self.assertEqual(parsed, expected)

    def test_custom_indentation(self):
        """Test getJson with custom indentation"""
        # Default indentation (2 spaces)
        json_2 = stdem.ExcelParser.getJson("tests/excel/example.xlsx", indent=2)
        self.assertIn("  ", json_2)  # 2 spaces

        # 4 spaces indentation
        json_4 = stdem.ExcelParser.getJson("tests/excel/example.xlsx", indent=4)
        self.assertIn("    ", json_4)  # 4 spaces

        # Compact (no indentation)
        json_0 = stdem.ExcelParser.getJson("tests/excel/example.xlsx", indent=0)

        # All should parse to same data
        data_2 = json.loads(json_2)
        data_4 = json.loads(json_4)
        data_0 = json.loads(json_0)

        self.assertEqual(data_2, data_4)
        self.assertEqual(data_2, data_0)


if __name__ == "__main__":
    unittest.main()
