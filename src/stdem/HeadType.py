from openpyxl.cell import Cell
from typing import Optional

from .TableException import (
    InvalidTypeNameError,
    InvalidHeaderFormatError,
    ChildAdditionError,
    UnexpectedDataError,
    TypeConversionError,
    InvalidIndexError,
)

type data = int | float | str | dict[str, data] | list[data] | None


class HeadType:
    def __init__(self, name: str, cell: Cell) -> None:
        self.name = name
        self.cell = cell
        self.column = cell.column - 2

    def addChild(self, child: "HeadType"):
        raise ChildAdditionError(
            self.cell, self.__class__.__name__, "This type does not support children"
        )

    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        if enable:
            if data[self.column].value is not None:
                return data[self.column].value
            else:
                return None
        elif data[self.column].value is not None:
            raise UnexpectedDataError(data[self.column], filename)

    def _validate_and_convert(
        self, cell: Cell, enable: bool, filename: Optional[str] = None
    ):
        """Helper method to validate cell data and handle conversion

        Args:
            cell: The cell to validate
            enable: Whether data is expected in this cell
            filename: Optional filename for error reporting

        Returns:
            Tuple of (should_process, cell_value) where should_process indicates
            if conversion should proceed and cell_value is the raw value

        Raises:
            UnexpectedDataError: If data found when enable=False
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
    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        should_process, value = self._validate_and_convert(
            data[self.column], enable, filename
        )
        if should_process:
            try:
                return int(value)
            except (ValueError, TypeError) as e:
                raise TypeConversionError(data[self.column], value, "int", e, filename)
        return None


class HeadString(HeadType):
    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
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
    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        should_process, value = self._validate_and_convert(
            data[self.column], enable, filename
        )
        if should_process:
            try:
                return float(value)
            except (ValueError, TypeError) as e:
                raise TypeConversionError(
                    data[self.column], value, "float", e, filename
                )
        return None


class HeadList(HeadType):
    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.key: HeadInt = None
        self.value: HeadType = None

    def addChild(self, child: HeadType):
        if self.key is None:
            if not isinstance(child, HeadInt):
                raise ChildAdditionError(
                    self.cell, "HeadList", "First child must be HeadInt (list index)"
                )
            self.key = child
        elif self.value is None:
            self.value = child
        else:
            raise ChildAdditionError(
                self.cell, "HeadList", "List can only have 2 children (index and value)"
            )

    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        if enable:
            self.data = []
        key = self.key.parseData(data, True, filename)
        if key is not None:
            if key != len(self.data):
                raise InvalidIndexError(
                    data[self.column], len(self.data), key, filename
                )
            self.data.append(self.value.parseData(data, True, filename))
        else:
            self.value.parseData(data, False, filename)
        if enable:
            return self.data


class HeadDict(HeadType):
    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.key: HeadString = None
        self.value: HeadType = None

    def addChild(self, child: HeadType):
        if self.key is None:
            if not isinstance(child, HeadString):
                raise ChildAdditionError(
                    self.cell, "HeadDict", "First child must be HeadString (dict key)"
                )
            self.key = child
        elif self.value is None:
            self.value = child
        else:
            raise ChildAdditionError(
                self.cell, "HeadDict", "Dict can only have 2 children (key and value)"
            )

    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        if enable:
            self.data = {}
        key = self.key.parseData(data, True, filename)
        if key is not None:
            self.data[key] = self.value.parseData(data, True, filename)
        else:
            self.value.parseData(data, False, filename)
        if enable:
            return self.data


class HeadClass(HeadType):
    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.children: list[HeadType] = []

    def addChild(self, child: "HeadType"):
        self.children.append(child)

    def parseData(
        self, data: list[Cell], enable: bool, filename: Optional[str] = None
    ) -> data:
        if enable:
            ret = {}
            for i in self.children:
                ret[i.name] = i.parseData(data, True, filename)
            return ret
        else:
            for i in self.children:
                i.parseData(data, False, filename)


typeDict: dict[str, type[HeadType]] = {
    "int": HeadInt,
    "string": HeadString,
    "float": HeadFloat,
    "list": HeadList,
    "dict": HeadDict,
    "class": HeadClass,
}


def headCreator(cell: Cell, filename: Optional[str] = None) -> HeadType:
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
        name, typeName = cell_value.split(":", 1)
    except ValueError as e:
        raise InvalidHeaderFormatError(
            cell, f"Invalid header format: {str(e)}", filename
        )

    if typeName not in typeDict:
        raise InvalidTypeNameError(cell, typeName, list(typeDict.keys()), filename)

    return typeDict[typeName](name, cell)
