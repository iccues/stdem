import argparse
import glob
import os

from . import ExcelParser


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("-o", type=str, default="json/")
    parser.add_argument("-dir", type=str, default=".")
    args = parser.parse_args()

    JSON_PATH = args.o
    EXCEL_PATH = args.dir

    parse_dir(EXCEL_PATH, JSON_PATH)


def parse_dir(excel_dir: str, json_dir: str):

    for filename in os.listdir(json_dir):
        os.remove(os.path.join(json_dir, filename))

    for filename in glob.glob("*.xlsx", root_dir=excel_dir):
        print(filename, end=":\t")
        excel_file = os.path.join(excel_dir, filename)
        json_file = os.path.join(json_dir, os.path.splitext(filename)[0] + ".json")
        parse_file(excel_file, json_file)


def parse_file(excel_file: str, json_file: str):
    try:
        jsonStr = ExcelParser.getJson(excel_file)
    except Exception as e:
        print(e)
    else:
        with open(json_file, "w") as file:
            file.write(jsonStr)
            print("Success!")
