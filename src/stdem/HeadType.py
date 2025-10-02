from openpyxl.cell import Cell
from typing import Optional

from .TableException import (
    InvalidTypeNameError,
    InvalidHeaderFormatError,
    ChildAdditionError,
    UnexpectedDataError,
    TypeConversionError,
    InvalidIndexError
)

type data = int | float | str | dict[str, data] | list[data] | None


class HeadType:

    def __init__(self, name: str, cell: Cell) -> None:
        self.name = name
        self.cell = cell
        self.colume = cell.column - 2

    def addChild(self, child: "HeadType"):
        raise ChildAdditionError(
            self.cell,
            self.__class__.__name__,
            "This type does not support children"
        )

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            if data[self.colume].value != None:
                return data[self.colume].value
            else:
                return None
        elif data[self.colume].value != None:
            raise UnexpectedDataError(data[self.colume], filename)

    def __repr__(self) -> str:
        return self.name



class HeadInt(HeadType):

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            if data[self.colume].value != None:
                try:
                    return int(data[self.colume].value)
                except (ValueError, TypeError) as e:
                    raise TypeConversionError(
                        data[self.colume],
                        data[self.colume].value,
                        "int",
                        e,
                        filename
                    )
            else:
                return None
        elif data[self.colume].value != None:
            raise UnexpectedDataError(data[self.colume], filename)


class HeadString(HeadType):

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            if data[self.colume].value != None:
                try:
                    return str(data[self.colume].value)
                except Exception as e:
                    raise TypeConversionError(
                        data[self.colume],
                        data[self.colume].value,
                        "string",
                        e,
                        filename
                    )
            else:
                return None
        elif data[self.colume].value != None:
            raise UnexpectedDataError(data[self.colume], filename)


class HeadFloat(HeadType):

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            if data[self.colume].value != None:
                try:
                    return float(data[self.colume].value)
                except (ValueError, TypeError) as e:
                    raise TypeConversionError(
                        data[self.colume],
                        data[self.colume].value,
                        "float",
                        e,
                        filename
                    )
            else:
                return None
        elif data[self.colume].value != None:
            raise UnexpectedDataError(data[self.colume], filename)



class HeadList(HeadType):

    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.key: HeadInt = None
        self.value: HeadType = None

    def addChild(self, child: HeadType):
        if self.key == None:
            if not isinstance(child, HeadInt):
                raise ChildAdditionError(
                    self.cell,
                    "HeadList",
                    "First child must be HeadInt (list index)"
                )
            self.key = child
        elif self.value == None:
            self.value = child
        else:
            raise ChildAdditionError(
                self.cell,
                "HeadList",
                "List can only have 2 children (index and value)"
            )

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            self.data = []
        key = self.key.parsetData(data, True, filename)
        if key != None:
            if key != len(self.data):
                raise InvalidIndexError(data[self.colume], len(self.data), key, filename)
            self.data.append(self.value.parsetData(data, True, filename))
        else:
            self.value.parsetData(data, False, filename)
        if enable:
            return self.data


class HeadDict(HeadType):

    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.key: HeadString = None
        self.value: HeadType = None

    def addChild(self, child: HeadType):
        if self.key == None:
            if not isinstance(child, HeadString):
                raise ChildAdditionError(
                    self.cell,
                    "HeadDict",
                    "First child must be HeadString (dict key)"
                )
            self.key = child
        elif self.value == None:
            self.value = child
        else:
            raise ChildAdditionError(
                self.cell,
                "HeadDict",
                "Dict can only have 2 children (key and value)"
            )

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            self.data = {}
        key = self.key.parsetData(data, True, filename)
        if key != None:
            self.data[key] = self.value.parsetData(data, True, filename)
        else:
            self.value.parsetData(data, False, filename)
        if enable:
            return self.data


class HeadClass(HeadType):

    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.children: list[HeadType] = []

    def addChild(self, child: "HeadType"):
        self.children.append(child)

    def parsetData(self, data: list[Cell], enable: bool, filename: Optional[str] = None) -> data:
        if enable:
            ret = {}
            for i in self.children:
                ret[i.name] = i.parsetData(data, True, filename)
            return ret
        else:
            for i in self.children:
                i.parsetData(data, False, filename)



typeDict: dict[str, type[HeadType]] = {
    "int" : HeadInt,
    "string" : HeadString,
    "float" : HeadFloat,
    "list" : HeadList,
    "dict" : HeadDict,
    "class" : HeadClass
}

def headCreater(cell: Cell, filename: Optional[str] = None) -> HeadType:
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
            cell,
            f"Header must be in format 'name:type', got: '{cell_value}'",
            filename
        )

    try:
        name, typeName = cell_value.split(":", 1)
    except ValueError as e:
        raise InvalidHeaderFormatError(
            cell,
            f"Invalid header format: {str(e)}",
            filename
        )

    if typeName not in typeDict:
        raise InvalidTypeNameError(
            cell,
            typeName,
            list(typeDict.keys()),
            filename
        )

    return typeDict[typeName](name, cell)
