import unittest
import json
from src import stdem


class TestParser(unittest.TestCase):

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

    def test_getJson_returns_valid_json(self):
        """Test that getJson returns valid JSON string"""
        json_str = stdem.ExcelParser.getJson("tests/excel/example.xlsx")

        # Should be valid JSON
        parsed = json.loads(json_str)
        self.assertIsInstance(parsed, dict)

        # Compare with expected
        with open("tests/json/example.json", "r", encoding="utf-8") as f:
            expected = json.load(f)
        self.assertEqual(parsed, expected)


if __name__ == "__main__":
    unittest.main()
