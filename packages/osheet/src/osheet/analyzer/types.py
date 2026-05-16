from __future__ import annotations
import datetime
from osheet.models import Cell, Sheet, Column, ColumnDtype


def _infer_dtype(cells: list[Cell]) -> ColumnDtype:
    if not cells:
        return ColumnDtype.UNKNOWN
    formula_count = sum(1 for c in cells if c.formula)
    if formula_count / len(cells) >= 0.5:
        return ColumnDtype.FORMULA
    values = [c.value for c in cells if c.value is not None]
    if not values:
        return ColumnDtype.UNKNOWN
    number_count = sum(1 for v in values if isinstance(v, (int, float)))
    date_count = sum(1 for v in values if isinstance(v, (datetime.date, datetime.datetime)))
    text_count = sum(1 for v in values if isinstance(v, str))
    total = len(values)
    if number_count / total >= 0.7:
        return ColumnDtype.NUMBER
    if date_count / total >= 0.7:
        return ColumnDtype.DATE
    if text_count / total >= 0.7:
        return ColumnDtype.TEXT
    return ColumnDtype.MIXED


def infer_column_types(sheet: Sheet) -> None:
    """Mutates each Column.dtype in place."""
    grid = {(c.row, c.col): c for c in sheet.cells}
    for table in sheet.tables:
        for col in table.columns:
            data_cells = [
                grid[(r, col.col_index)]
                for r in range(table.first_data_row, table.last_data_row + 1)
                if (r, col.col_index) in grid
            ]
            col.dtype = _infer_dtype(data_cells)
