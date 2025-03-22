from openpyxl.cell import Cell


class TableException(Exception):
    def __init__(self, cell: Cell, message: str) -> None:
        super().__init__(f"{cell.coordinate}: {message}")

class TableHeadError(TableException):
    pass

class TableDataError(TableException):
    pass
