# stdem

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/iccues/stdem/blob/main/LICENSE)
[![PyPI - Version](https://img.shields.io/pypi/v/stdem)](https://pypi.org/project/stdem/)

A powerful tool for converting Excel spreadsheets into JSON data with complex hierarchical structures.

[中文文档](README.md) | English

## ✨ Features

- 🔄 **Complex Data Structure Support** - Supports nested objects, lists, dictionaries, and other complex hierarchical structures
- 📊 **Type Safe** - Built-in type validation and conversion (int, float, string, list, dict, object)
- 🎯 **Detailed Error Messages** - Precisely locates error cells and reasons
- 🚀 **Batch Processing** - Process entire directories of Excel files at once
- 🔍 **Formatted Output** - Generates formatted JSON files for easy reading

## 📦 Installation

Install using pip:

```bash
pip install stdem
```

Or install using uv:

```bash
uv pip install stdem
```

## 🚀 Quick Start

### Command Line Usage

stdem provides a modern subcommand interface:

#### Convert Files/Directories

```bash
# Convert a single file
stdem convert input.xlsx -o output.json

# Convert entire directory
stdem convert excel_dir/ -o json_dir/

# Custom JSON indentation
stdem convert data/ -o output/ --indent 4

# Keep existing files in output directory (don't clear)
stdem convert data/ -o output/ --no-clear

# Silent mode (show errors only)
stdem convert data/ -o output/ -q

# Enable verbose error output
stdem convert data/ -o output/ -v
```

#### Validate File Format

```bash
# Validate if Excel file conforms to stdem format
stdem validate config.xlsx

# Show detailed validation info and data preview
stdem validate config.xlsx -v
```

#### Other Commands

```bash
# View help
stdem --help
stdem convert --help
stdem validate --help

# View version
stdem --version
```

### Command Line Arguments

#### `stdem convert` Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `input` | Input Excel file or directory (required) | - |
| `-o, --output` | Output JSON file or directory (required) | - |
| `-i, --indent` | JSON indentation spaces | 2 |
| `--no-clear` | Don't clear output directory | false |
| `-q, --quiet` | Silent mode (show errors only) | false |
| `-v, --verbose` | Verbose error output | false |

#### `stdem validate` Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `file` | Excel file to validate (required) | - |
| `-v, --verbose` | Show detailed validation info | false |

**⚠️ Note:**

- By default, all `.json` files in the output directory will be cleared when converting directories
- Use `--no-clear` parameter to keep existing files
- Single file conversion automatically creates output directory if it doesn't exist

### Python API Usage

```python
from stdem import excel_parser

# Parse a single file to Python object
data = excel_parser.get_data("example.xlsx")

# Parse a single file to JSON string
json_str = excel_parser.get_json("example.xlsx")

# Batch process directory
from stdem import main
success, failed = main.parse_dir("excel/", "json/")
print(f"Success: {success}, Failed: {failed}")
```

## 📋 Excel Format Specification

### Basic Format

Excel files must follow this format:

1. **First row, first column** must be the `#head` marker
2. **First row, other columns** define field names and types in format `fieldname:type`
3. **Header rows** can have multiple rows to define complex nested structures
4. **Data rows** must start with `#data` marker
5. **Comment rows** starting with `#` are ignored

### Supported Data Types

| Type | Description | Example |
|------|-------------|---------|
| `int` | Integer | `age:int` |
| `float` | Float | `price:float` |
| `string` | String | `name:string` |
| `list` | List (requires two sub-columns: index and value) | `items:list` |
| `dict` | Dictionary (requires two sub-columns: key and value) | `config:dict` |
| `object` | Nested object | `player:object` |
| ~~`class`~~ | **Deprecated**, use `object` instead | ~~`player:class`~~ |

> **⚠️ Breaking Change (v0.3.0)**: The `class` type has been renamed to `object`. For backward compatibility, `class` is still supported but deprecated. It's recommended to use `object` in new projects.

### Example: Simple Table

![Excel format example](https://raw.githubusercontent.com/iccues/stdem/main/docs/image/example.png)

Converts to:

```json
{
  "Nyxra": {
    "hp": 10000,
    "attack": 200.0,
    "skills": ["Shadowstep", "Twilight Veil", "Void Requiem"]
  },
  "Orin": {
    "hp": 15000,
    "attack": 100.0,
    "skills": ["Mana Surge", "Celestial Wrath"]
  }
}
```

## 🔍 Error Handling

stdem provides detailed error messages to help quickly locate issues:

### File Related Errors

- `FileNotFoundError` - File doesn't exist
- `InvalidFileFormatError` - Unsupported file format (must be .xlsx or .xlsm)
- `EmptyFileError` - File is empty

### Header Related Errors

- `MissingHeaderMarkerError` - Missing `#head` marker
- `InvalidHeaderFormatError` - Invalid header format (should be `name:type`)
- `InvalidTypeNameError` - Invalid type name
- `ChildAdditionError` - Child node addition error

### Data Related Errors

- `MissingDataMarkerError` - Missing `#data` marker
- `UnexpectedDataError` - Data found in disabled cell
- `TypeConversionError` - Type conversion failed
- `InvalidIndexError` - List index error

Error example:

```bash
$ stdem convert excel/ -o json/ -v

example.xlsx:   [OK] Success!
invalid.xlsx:   [ERROR] File: invalid.xlsx | Cell: B1 | Invalid type 'wrong'. Valid types: int, string, float, list, dict, object

[DONE] Processing complete: 1 succeeded, 1 failed
```

Validation example:

```bash
$ stdem validate tests/excel/example.xlsx -v
[OK] example.xlsx is valid!

Data structure preview:
{
  "Nyxra": {
    "hp": 10000,
    "attack": 200.0,
    "skills": ["Shadowstep", "Twilight Veil", "Void Requiem"]
  },
  ...
}
```

## 🛠️ Development

### Install Dependencies

```bash
# Using uv
uv sync

# Or using pip
pip install -e ".[dev]"
```

### Run Tests

```bash
# Run all tests
python -m unittest discover tests -v

# Run specific test modules
python -m unittest tests.test_parsing -v
python -m unittest tests.test_errors -v
```

### Test Structure

Test files are organized by functionality:

- `tests/test_base.py` - Base utility class and fixtures tests
- `tests/test_parsing.py` - Excel parsing and JSON formatting tests
- `tests/test_errors.py` - Error handling and validation tests

### Project Structure

```text
stdem/
├── src/stdem/
│   ├── __init__.py
│   ├── __main__.py
│   ├── ExcelParser.py      # Excel parsing core
│   ├── HeadType.py          # Type definition and conversion
│   ├── TableException.py    # Exception definitions
│   └── Main.py              # CLI entry point
├── tests/
│   ├── excel/               # Test Excel files
│   ├── json/                # Expected JSON results
│   ├── test_base.py         # Test base utilities
│   ├── test_parsing.py      # Parsing functionality tests
│   └── test_errors.py       # Error handling tests
├── pyproject.toml
└── README.md
```

## 📝 Use Cases

- **Game Development** - Convert Excel configurations from planners into game data
- **Data Migration** - Import Excel data into databases or other systems
- **Configuration Management** - Manage complex configuration files with Excel
- **Data Exchange** - Convert between Excel and JSON formats

## 🤝 Contributing

Issues and Pull Requests are welcome!

## 📄 License

This project is open source under the [MIT License](https://github.com/iccues/stdem/blob/main/LICENSE).

## 🔗 Links

- [PyPI](https://pypi.org/project/stdem/)
- [GitHub](https://github.com/iccues/stdem)
- [Issue Tracker](https://github.com/iccues/stdem/issues)

---

Made with ❤️ by [iccues](https://github.com/iccues)
