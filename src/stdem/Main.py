import argparse
import glob
import os
import sys
from typing import Tuple
from pathlib import Path

from . import ExcelParser
from .TableException import TableException


def main():
    """Main entry point for the stdem CLI"""
    parser = argparse.ArgumentParser(
        prog="stdem",
        description="Convert Excel tables to JSON format with complex hierarchical structures",
        epilog="Examples:\n"
               "  stdem convert input.xlsx -o output.json          # Convert single file\n"
               "  stdem convert excel_dir/ -o json_dir/            # Convert directory\n"
               "  stdem convert data/ -o output/ --indent 4        # Custom indentation\n"
               "  stdem validate config.xlsx                       # Validate Excel file\n"
               "  stdem --version                                  # Show version",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s 0.1.1"
    )

    # Create subparsers for different commands
    subparsers = parser.add_subparsers(
        dest="command",
        help="Available commands"
    )

    # Convert command (default)
    convert_parser = subparsers.add_parser(
        "convert",
        help="Convert Excel file(s) to JSON",
        description="Convert Excel tables to JSON format"
    )
    convert_parser.add_argument(
        "input",
        type=str,
        help="Input Excel file or directory"
    )
    convert_parser.add_argument(
        "-o", "--output",
        type=str,
        required=True,
        metavar="PATH",
        help="Output JSON file or directory (required)"
    )
    convert_parser.add_argument(
        "-i", "--indent",
        type=int,
        default=2,
        metavar="N",
        help="JSON indentation spaces (default: 2)"
    )
    convert_parser.add_argument(
        "--no-clear",
        action="store_true",
        help="Don't clear output directory before conversion"
    )
    convert_parser.add_argument(
        "-q", "--quiet",
        action="store_true",
        help="Suppress output (only show errors)"
    )
    convert_parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose error output"
    )

    # Validate command
    validate_parser = subparsers.add_parser(
        "validate",
        help="Validate Excel file format",
        description="Check if Excel file follows stdem format specification"
    )
    validate_parser.add_argument(
        "file",
        type=str,
        help="Excel file to validate"
    )
    validate_parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Show detailed validation information"
    )

    args = parser.parse_args()

    # Require a command
    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Execute command
    if args.command == "convert":
        return convert_command(args)
    elif args.command == "validate":
        return validate_command(args)


def convert_command(args):
    """Handle convert subcommand"""
    input_path = Path(args.input)
    output_path = Path(args.output)

    try:
        # Single file conversion
        if input_path.is_file():
            if not output_path.suffix:
                output_path = output_path / f"{input_path.stem}.json"

            if not args.quiet:
                print(f"Converting {input_path} -> {output_path}")

            success = parse_file(
                str(input_path),
                str(output_path),
                verbose=args.verbose,
                indent=args.indent,
                quiet=args.quiet
            )
            sys.exit(0 if success else 1)

        # Directory conversion
        elif input_path.is_dir():
            stats = parse_dir(
                str(input_path),
                str(output_path),
                verbose=args.verbose,
                indent=args.indent,
                clear_output=not args.no_clear,
                quiet=args.quiet
            )
            if not args.quiet:
                print(f"\n[DONE] Processing complete: {stats[0]} succeeded, {stats[1]} failed")
            sys.exit(0 if stats[1] == 0 else 1)

        else:
            print(f"Error: Input path does not exist: {input_path}", file=sys.stderr)
            sys.exit(1)

    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


def validate_command(args):
    """Handle validate subcommand"""
    file_path = Path(args.file)

    if not file_path.exists():
        print(f"[ERROR] File not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    try:
        data = ExcelParser.getData(str(file_path))
        print(f"[OK] {file_path.name} is valid!")

        if args.verbose:
            import json
            print(f"\nData structure preview:")
            preview = json.dumps(data, indent=2, ensure_ascii=False)
            lines = preview.split('\n')
            print('\n'.join(lines[:20]))
            if len(lines) > 20:
                print(f"... ({len(lines) - 20} more lines)")

        sys.exit(0)
    except TableException as e:
        print(f"[ERROR] Validation failed: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] Unexpected error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


def parse_dir(
    excel_dir: str,
    json_dir: str,
    verbose: bool = False,
    indent: int = 2,
    clear_output: bool = True,
    quiet: bool = False
) -> Tuple[int, int]:
    """Parse all Excel files in a directory

    Args:
        excel_dir: Directory containing Excel files
        json_dir: Output directory for JSON files
        verbose: Enable verbose error output
        indent: JSON indentation level
        clear_output: Clear output directory before processing
        quiet: Suppress non-error output

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
    if clear_output:
        for filename in os.listdir(json_dir):
            file_path = os.path.join(json_dir, filename)
            if os.path.isfile(file_path) and filename.endswith('.json'):
                os.remove(file_path)

    # Process Excel files
    success_count = 0
    failure_count = 0

    excel_files = list(glob.glob("*.xlsx", root_dir=excel_dir))

    if not excel_files:
        if not quiet:
            print(f"Warning: No Excel files found in {excel_dir}")
        return (0, 0)

    for filename in excel_files:
        if not quiet:
            print(f"{filename}:\t", end="")
        excel_file = os.path.join(excel_dir, filename)
        json_file = os.path.join(json_dir, os.path.splitext(filename)[0] + ".json")

        if parse_file(excel_file, json_file, verbose, indent, quiet):
            success_count += 1
        else:
            failure_count += 1

    return (success_count, failure_count)


def parse_file(
    excel_file: str,
    json_file: str,
    verbose: bool = False,
    indent: int = 2,
    quiet: bool = False
) -> bool:
    """Parse a single Excel file to JSON

    Args:
        excel_file: Path to Excel file
        json_file: Path to output JSON file
        verbose: Enable verbose error output
        indent: JSON indentation level
        quiet: Suppress non-error output

    Returns:
        True if successful, False otherwise
    """
    try:
        jsonStr = ExcelParser.getJson(excel_file, indent=indent)

        # Ensure output directory exists
        os.makedirs(os.path.dirname(json_file) or ".", exist_ok=True)

        with open(json_file, "w", encoding="utf-8") as file:
            file.write(jsonStr)

        if not quiet:
            print("[OK] Success!")
        return True
    except TableException as e:
        # Handle our custom exceptions with detailed error messages
        print(f"[ERROR] {e}", file=sys.stderr)
        if verbose:
            import traceback
            traceback.print_exc()
        return False
    except Exception as e:
        # Handle unexpected errors
        print(f"[ERROR] Unexpected error: {type(e).__name__}: {e}", file=sys.stderr)
        if verbose:
            import traceback
            traceback.print_exc()
        return False
