import unittest
from src import stdem

class TestParser(unittest.TestCase):
    def test_parser(self):
        print(stdem.ExcelParser.getJson("tests/excel/Table.xlsx"))

