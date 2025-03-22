import unittest
from src import stdem

stdem.Main.parse_dir("tests/json", "tests/excel")

class TestMain(unittest.TestCase):
    pass
