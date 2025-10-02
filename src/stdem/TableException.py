from openpyxl.cell import Cell
from typing import Optional


class TableException(Exception):
    """Base exception for all table-related errors"""
    def __init__(self, message: str, cell: Optional[Cell] = None, filename: Optional[str] = None) -> None:
        self.cell = cell
        self.filename = filename

        error_parts = []
        if filename:
            error_parts.append(f"File: {filename}")
        if cell:
            error_parts.append(f"Cell: {cell.coordinate}")
        error_parts.append(message)

        super().__init__(" | ".join(error_parts))


# File-related errors
class FileError(TableException):
    """Base class for file-related errors"""
    pass


class FileNotFoundError(FileError):
    """File does not exist"""
    def __init__(self, filename: str) -> None:
        super().__init__(f"File not found: {filename}", filename=filename)


class InvalidFileFormatError(FileError):
    """File format is invalid or not supported"""
    def __init__(self, filename: str, expected_format: str = "xlsx") -> None:
        super().__init__(
            f"Invalid file format. Expected: {expected_format}",
            filename=filename
        )


class EmptyFileError(FileError):
    """File is empty or has no valid data"""
    def __init__(self, filename: str) -> None:
        super().__init__("File is empty or contains no data", filename=filename)


# Header-related errors
class TableHeadError(TableException):
    """Base class for table header errors"""
    pass


class MissingHeaderMarkerError(TableHeadError):
    """First cell must be '#head'"""
    def __init__(self, cell: Cell, actual_value: str, filename: Optional[str] = None) -> None:
        super().__init__(
            f"Expected '#head' marker, found: '{actual_value}'",
            cell=cell,
            filename=filename
        )


class InvalidHeaderFormatError(TableHeadError):
    """Header format is incorrect"""
    def __init__(self, cell: Cell, message: str, filename: Optional[str] = None) -> None:
        super().__init__(f"Invalid header format: {message}", cell=cell, filename=filename)


class InvalidTypeNameError(TableHeadError):
    """Type name in header is invalid"""
    def __init__(self, cell: Cell, typename: str, valid_types: list[str], filename: Optional[str] = None) -> None:
        super().__init__(
            f"Invalid type '{typename}'. Valid types: {', '.join(valid_types)}",
            cell=cell,
            filename=filename
        )


class ChildAdditionError(TableHeadError):
    """Cannot add child to this header type"""
    def __init__(self, cell: Cell, parent_type: str, reason: str, filename: Optional[str] = None) -> None:
        super().__init__(
            f"Cannot add child to {parent_type}: {reason}",
            cell=cell,
            filename=filename
        )


# Data-related errors
class TableDataError(TableException):
    """Base class for table data errors"""
    pass


class MissingDataMarkerError(TableDataError):
    """No '#data' marker found in file"""
    def __init__(self, filename: Optional[str] = None) -> None:
        super().__init__(
            "No '#data' marker found. Table must have a data section.",
            filename=filename
        )


class UnexpectedDataError(TableDataError):
    """Data found in unexpected location"""
    def __init__(self, cell: Cell, filename: Optional[str] = None) -> None:
        super().__init__(
            "Unexpected data encountered in disabled cell",
            cell=cell,
            filename=filename
        )


class TypeConversionError(TableDataError):
    """Failed to convert data to expected type"""
    def __init__(self, cell: Cell, value: any, target_type: str, error: Exception, filename: Optional[str] = None) -> None:
        super().__init__(
            f"Cannot convert '{value}' to {target_type}: {str(error)}",
            cell=cell,
            filename=filename
        )


class InvalidIndexError(TableDataError):
    """List index is invalid or out of sequence"""
    def __init__(self, cell: Cell, expected: int, actual: int, filename: Optional[str] = None) -> None:
        super().__init__(
            f"Invalid list index. Expected: {expected}, Got: {actual}",
            cell=cell,
            filename=filename
        )


class MissingRequiredDataError(TableDataError):
    """Required data field is missing"""
    def __init__(self, cell: Cell, field_name: str, filename: Optional[str] = None) -> None:
        super().__init__(
            f"Missing required data for field: {field_name}",
            cell=cell,
            filename=filename
        )
