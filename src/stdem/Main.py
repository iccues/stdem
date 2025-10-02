import argparse
import glob
import os
import sys
from typing import Tuple

from . import ExcelParser
from .TableException import TableException


def main():
    """Main entry point for the stdem CLI"""
    parser = argparse.ArgumentParser(
        description="Convert Excel tables to JSON format",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "-o",
        type=str,
        default="json/",
        help="Output directory for JSON files (default: json/)"
    )
    parser.add_argument(
        "-dir",
        type=str,
        default=".",
        help="Input directory containing Excel files (default: current directory)"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose error output"
    )

    args = parser.parse_args()

    try:
        stats = parse_dir(args.dir, args.o, verbose=args.verbose)
        print(f"\nProcessing complete: {stats[0]} succeeded, {stats[1]} failed")
        sys.exit(0 if stats[1] == 0 else 1)
    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


def parse_dir(excel_dir: str, json_dir: str, verbose: bool = False) -> Tuple[int, int]:
    """Parse all Excel files in a directory

    Args:
        excel_dir: Directory containing Excel files
        json_dir: Output directory for JSON files
        verbose: Enable verbose error output

    Returns:
        Tuple of (success_count, failure_count)

    Raises:
        FileNotFoundError: If excel_dir doesn't exist
        NotADirectoryError: If excel_dir is not a directory
    """
    # Validate input directory
    if not os.path.exists(excel_dir):
        raise FileNotFoundError(f"Input directory not found: {excel_dir}")

    if not os.path.isdir(excel_dir):
        raise NotADirectoryError(f"Not a directory: {excel_dir}")

    # Create output directory if needed
    os.makedirs(json_dir, exist_ok=True)

    # Clear existing JSON files
    for filename in os.listdir(json_dir):
        file_path = os.path.join(json_dir, filename)
        if os.path.isfile(file_path) and filename.endswith('.json'):
            os.remove(file_path)

    # Process Excel files
    success_count = 0
    failure_count = 0

    excel_files = list(glob.glob("*.xlsx", root_dir=excel_dir))

    if not excel_files:
        print(f"Warning: No Excel files found in {excel_dir}")
        return (0, 0)

    for filename in excel_files:
        print(f"{filename}:\t", end="")
        excel_file = os.path.join(excel_dir, filename)
        json_file = os.path.join(json_dir, os.path.splitext(filename)[0] + ".json")

        if parse_file(excel_file, json_file, verbose):
            success_count += 1
        else:
            failure_count += 1

    return (success_count, failure_count)


def parse_file(excel_file: str, json_file: str, verbose: bool = False) -> bool:
    """Parse a single Excel file to JSON

    Args:
        excel_file: Path to Excel file
        json_file: Path to output JSON file
        verbose: Enable verbose error output

    Returns:
        True if successful, False otherwise
    """
    try:
        jsonStr = ExcelParser.getJson(excel_file)
        with open(json_file, "w", encoding="utf-8") as file:
            file.write(jsonStr)
        print("Success!")
        return True
    except TableException as e:
        # Handle our custom exceptions with detailed error messages
        print(f"Error: {e}")
        if verbose:
            import traceback
            traceback.print_exc()
        return False
    except Exception as e:
        # Handle unexpected errors
        print(f"Unexpected error: {type(e).__name__}: {e}")
        if verbose:
            import traceback
            traceback.print_exc()
        return False
