import json
from src import stdem


def test_get_python_object(test_json_dir):
    result = stdem.excel_parser.get_data("tests/excel/example.xlsx")
    with open(test_json_dir / "example.json", "r", encoding="utf-8") as f:
        expected = json.load(f)

    assert result == expected

def test_get_json_string(test_json_dir):
    result = stdem.excel_parser.get_json("tests/excel/example.xlsx")
    with open(test_json_dir / "example.json", "r", encoding="utf-8") as f:
        excepted = f.read()

    assert result == excepted
