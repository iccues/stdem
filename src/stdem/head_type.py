from openpyxl.cell import Cell
from typing import Optional

from .constants import VALID_TYPES
from .exceptions import (
    InvalidTypeNameError,
    InvalidHeaderFormatError,
    ChildAdditionError,
    UnexpectedDataError,
    TypeConversionError,
    InvalidIndexError,
)

# Data type definition: supports complex nested structures
type data = int | float | str | dict[str, data] | list[data] | None


class HeadType:
    """Base class for header types, defining the common interface for all header types

    Header types are used to define the data type and structure of Excel columns, supporting:
    - Basic types: int, float, string
    - Complex types: list, dict, object

    Attributes:
        name: Field name
        cell: Corresponding Excel cell
        column: Data column index (relative to header start position)
    """

    def __init__(self, name: str, cell: Cell) -> None:
        """Initialize header type

        Args:
            name: Field name
            cell: Excel cell reference
        """
        self.name = name
        self.cell = cell
        # Subtract 2 from column index: column 0 is marker (#head/#data), column 1+ is actual data
        self.column = cell.column - 2

    def add_child(self, child: "HeadType") -> None:
        """Add child header (for complex types)

        Base class does not support child headers, only HeadList, HeadDict, HeadObject support it

        Args:
            child: Child header object

        Raises:
            ChildAdditionError: Current type does not support adding child nodes
        """
        raise ChildAdditionError(
            self.cell, self.__class__.__name__, "This type does not support children"
        )

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse cell data

        Args:
            data: List of cells in the data row
            enable: If True, extract and return data from this column.
                   If False, validate that this column is empty (used for columns under
                   merged headers that shouldn't have data in this row)
            filename: Optional filename for error reporting

        Returns:
            Parsed data value if enable=True and cell has data, otherwise None

        Raises:
            UnexpectedDataError: When enable=False but cell contains data
        """
        if enable:
            # Extract and return data from this column
            if data[self.column].value is not None:
                return data[self.column].value
            else:
                return None
        elif data[self.column].value is not None:
            # Validation mode: cell should be empty but has data
            raise UnexpectedDataError(data[self.column], filename)
        # Validation mode: cell is empty as expected
        return None

    def _validate_and_convert(
        self, cell: Cell, enable: bool, filename: Optional[str] = None
    ) -> tuple[bool, any]:
        """Helper method to validate cell data and handle conversion

        Args:
            cell: Cell to validate
            enable: If True, expect and extract data. If False, validate cell is empty
            filename: Optional filename for error reporting

        Returns:
            Tuple of (should_process, cell_value) where:
            - should_process: True if value should be converted, False otherwise
            - cell_value: The raw cell value or None

        Raises:
            UnexpectedDataError: When enable=False but data is found
        """
        if enable:
            if cell.value is not None:
                return True, cell.value
            else:
                return False, None
        else:
            if cell.value is not None:
                raise UnexpectedDataError(cell, filename)
            return False, None

    def __repr__(self) -> str:
        return self.name


class HeadInt(HeadType):
    """Integer type header"""

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse integer data"""
        should_process, value = self._validate_and_convert(
            data[self.column], enable, filename
        )
        if should_process:
            try:
                return int(value)
            except Exception as e:
                raise TypeConversionError(data[self.column], value, "int", e, filename)
        return None


class HeadString(HeadType):
    """String type header"""

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse string data"""
        should_process, value = self._validate_and_convert(
            data[self.column], enable, filename
        )
        if should_process:
            try:
                return str(value)
            except Exception as e:
                raise TypeConversionError(
                    data[self.column], value, "string", e, filename
                )
        return None


class HeadFloat(HeadType):
    """Float type header"""

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse float data"""
        should_process, value = self._validate_and_convert(
            data[self.column], enable, filename
        )
        if should_process:
            try:
                return float(value)
            except Exception as e:
                raise TypeConversionError(
                    data[self.column], value, "float", e, filename
                )
        return None


class HeadList(HeadType):
    """List type header

    List requires two child columns:
    1. Index column (must be HeadInt) - used to specify element order
    2. Value column (any type) - actual value of list element

    Attributes:
        key: Header of index column (HeadInt type)
        value: Header of value column (any HeadType type)
        data: Accumulated list data
    """

    def __init__(self, name: str, cell: Cell) -> None:
        """Initialize list type header"""
        super().__init__(name, cell)
        self.key: HeadInt = None  # Index column
        self.value: HeadType = None  # Value column

    def add_child(self, child: HeadType) -> None:
        """Add child header

        Must be added in order: index column first, then value column

        Args:
            child: Child header object

        Raises:
            ChildAdditionError: Child header type error or count exceeds limit
        """
        if self.key is None:
            # First child node must be integer type (used as index)
            if not isinstance(child, HeadInt):
                raise ChildAdditionError(
                    self.cell, "HeadList", "First child must be HeadInt (list index)"
                )
            self.key = child
        elif self.value is None:
            # Second child node is the value type
            self.value = child
        else:
            # List can only have two child nodes
            raise ChildAdditionError(
                self.cell, "HeadList", "List can only have 2 children (index and value)"
            )

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse list data

        List data is built by accumulating multiple rows, each row adds one element
        Index must be consecutive (0, 1, 2, ...)
        """
        if enable:
            # First enable: initialize empty list
            self.data = []

        # Parse index value
        key = self.key.parse_data(data, True, filename)
        if key is not None:
            # Has index value: validate index continuity and add element
            if key != len(self.data):
                raise InvalidIndexError(
                    data[self.column], len(self.data), key, filename
                )
            self.data.append(self.value.parse_data(data, True, filename))
        else:
            # No index value: only validate that value column should not have data
            self.value.parse_data(data, False, filename)

        if enable:
            return self.data
        return None


class HeadDict(HeadType):
    """Dictionary type header

    Dictionary requires two child columns:
    1. Key column (must be HeadString) - dictionary key
    2. Value column (any type) - dictionary value

    Attributes:
        key: Header of key column (HeadString type)
        value: Header of value column (any HeadType type)
        data: Accumulated dictionary data
    """

    def __init__(self, name: str, cell: Cell) -> None:
        """Initialize dictionary type header"""
        super().__init__(name, cell)
        self.key: HeadString = None  # Key column
        self.value: HeadType = None  # Value column

    def add_child(self, child: HeadType) -> None:
        """Add child header

        Must be added in order: key column first, then value column

        Args:
            child: Child header object

        Raises:
            ChildAdditionError: Child header type error or count exceeds limit
        """
        if self.key is None:
            # First child node must be string type (used as key)
            if not isinstance(child, HeadString):
                raise ChildAdditionError(
                    self.cell, "HeadDict", "First child must be HeadString (dict key)"
                )
            self.key = child
        elif self.value is None:
            # Second child node is the value type
            self.value = child
        else:
            # Dictionary can only have two child nodes
            raise ChildAdditionError(
                self.cell, "HeadDict", "Dict can only have 2 children (key and value)"
            )

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse dictionary data

        Dictionary data is built by accumulating multiple rows, each row adds one key-value pair
        """
        if enable:
            # First enable: initialize empty dictionary
            self.data = {}

        # Parse key
        key = self.key.parse_data(data, True, filename)
        if key is not None:
            # Has key: add key-value pair
            self.data[key] = self.value.parse_data(data, True, filename)
        else:
            # No key: only validate that value column should not have data
            self.value.parse_data(data, False, filename)

        if enable:
            return self.data
        return None


class HeadObject(HeadType):
    """Object type header

    Object type is used to represent nested object structures, can contain multiple child fields
    Each child field can be any type (including other object types, forming multi-level nesting)

    Attributes:
        children: List of child fields, each child field is a HeadType object
    """

    def __init__(self, name: str, cell: Cell) -> None:
        """Initialize object type header"""
        super().__init__(name, cell)
        self.children: list[HeadType] = []

    def add_child(self, child: "HeadType") -> None:
        """Add child field

        Object type can add any number of child fields

        Args:
            child: Header object of the child field
        """
        self.children.append(child)

    def parse_data(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        """Parse object data

        Combine all child field data into a dictionary object
        """
        if enable:
            # Enable mode: parse all child fields and build dictionary
            ret = {}
            for i in self.children:
                ret[i.name] = i.parse_data(data, True, filename)
            return ret
        else:
            # Disable mode: only validate all child fields
            for i in self.children:
                i.parse_data(data, False, filename)
        return None


# Type name to class mapping table, used to create corresponding header objects based on string type names
TYPE_DICT: dict[str, type[HeadType]] = {
    "int": HeadInt,
    "string": HeadString,
    "float": HeadFloat,
    "list": HeadList,
    "dict": HeadDict,
    "object": HeadObject,
    "class": HeadObject,  # Deprecated: kept for backward compatibility, use 'object' instead
}


def head_creator(cell: Cell, filename: Optional[str] = None) -> HeadType:
    """Create a HeadType instance from a cell

    Args:
        cell: Cell containing header definition (format: "name:type")
        filename: Optional filename for error reporting

    Returns:
        HeadType instance of appropriate type

    Raises:
        InvalidHeaderFormatError: If cell format is invalid
        InvalidTypeNameError: If type name is not recognized
    """
    cell_value = str(cell.value) if cell.value else ""

    if ":" not in cell_value:
        raise InvalidHeaderFormatError(
            cell, f"Header must be in format 'name:type', got: '{cell_value}'", filename
        )

    try:
        name, type_name = cell_value.split(":", 1)
    except ValueError as e:
        raise InvalidHeaderFormatError(
            cell, f"Invalid header format: {str(e)}", filename
        )

    if type_name not in VALID_TYPES:
        raise InvalidTypeNameError(cell, type_name, list(VALID_TYPES), filename)

    return TYPE_DICT[type_name](name, cell)
