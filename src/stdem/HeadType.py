from openpyxl.cell import Cell

from .TableException import *

type data = int | float | str | dict[str, data] | list[data] | None

class HeadType:

    def __init__(self, name: str, cell: Cell) -> None:
        self.name = name
        self.cell = cell
        self.colume = cell.column - 2

    def addChild(self, child: "HeadType"):
        raise TableHeadError(self.cell, "unable add child on this cell.")

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            if data[self.colume].value != None:
                return data[self.colume].value
            else:
                return None
        elif data[self.colume].value != None:
            raise TableDataError(data[self.colume], "")

    def __repr__(self) -> str:
        return self.name



class HeadInt(HeadType):

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            if data[self.colume].value != None:
                return int(data[self.colume].value)
            else:
                return None
        elif data[self.colume].value != None:
            raise TableDataError(data[self.colume], "")


class HeadString(HeadType):

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            if data[self.colume].value != None:
                return str(data[self.colume].value)
            else:
                return None
        elif data[self.colume].value != None:
            raise TableDataError(data[self.colume], "")


class HeadFloat(HeadType):

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            if data[self.colume].value != None:
                return float(data[self.colume].value)
            else:
                return None
        elif data[self.colume].value != None:
            raise TableDataError(data[self.colume], "")



class HeadList(HeadType):

    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.key: HeadInt = None
        self.value: HeadType = None

    def addChild(self, child: HeadType):
        if self.key == None:
            if not isinstance(child, HeadInt):
                raise TableHeadError(self.cell, "unable addchild on this type")
            self.key = child
        elif self.value == None:
            self.value = child
        else:
            raise TableHeadError(self.cell, "unable add more child on this cell")

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            self.data = []
        key = self.key.parsetData(data, True)
        if key != None:
            if key != len(self.data):
                raise TableDataError(data[self.colume], "index number type error")
            self.data.append(self.value.parsetData(data, True))
        else:
            self.value.parsetData(data, False)
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
                raise TableHeadError(self.cell, "unable addchild on this type")
            self.key = child
        elif self.value == None:
            self.value = child
        else:
            raise TableHeadError(self.cell, "unable add more child on this cell")

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            self.data = {}
        key = self.key.parsetData(data, True)
        if key != None:
            self.data[key] = self.value.parsetData(data, True)
        else:
            self.value.parsetData(data, False)
        if enable:
            return self.data


class HeadClass(HeadType):

    def __init__(self, name: str, cell: Cell) -> None:
        super().__init__(name, cell)
        self.children: list[HeadType] = []

    def addChild(self, child: "HeadType"):
        self.children.append(child)

    def parsetData(self, data: list[Cell], enable: bool) -> data:
        if enable:
            ret = {}
            for i in self.children:
                ret[i.name] = i.parsetData(data, True)
            return ret
        else:
            for i in self.children:
                i.parsetData(data, False)



typeDict: dict[str, type[HeadType]] = {
    "int" : HeadInt,
    "string" : HeadString,
    "float" : HeadFloat,
    "list" : HeadList,
    "dict" : HeadDict,
    "class" : HeadClass
}

def headCreater(cell: Cell) -> HeadType:
    try:
        name, typeName = str(cell.value).split(":")
        ret = typeDict[typeName](name, cell)
    except:
        raise TableHeadError(cell, "typename error")
    return ret
