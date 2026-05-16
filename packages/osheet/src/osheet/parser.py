from __future__ import annotations
import io
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell as OxlCell
from osheet.models import Cell, CellRole, Sheet, Workbook, Manifest


def _fill_color(cell: OxlCell) -> str | None:
    try:
        fill = cell.fill
        if fill and fill.fill_type not in (None, "none"):
            fg = fill.fgColor
            if fg and fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF"):
                return fg.rgb
    except Exception:
        pass
    return None


def parse_xlsx(data: bytes) -> Workbook:
    wb_ox = load_workbook(io.BytesIO(data), data_only=False)
    sheets: list[Sheet] = []

    for ws in wb_ox.worksheets:
        cells: list[Cell] = []
        for row in ws.iter_rows():
            for ox_cell in row:
                if ox_cell.value is None:
                    continue
                raw_val = ox_cell.value
                formula: str | None = None
                value = raw_val
                if isinstance(raw_val, str) and raw_val.startswith("="):
                    formula = raw_val
                    value = None

                cells.append(Cell(
                    stable_id=f"{ws.title}.{ox_cell.column_letter}{ox_cell.row}",
                    role=CellRole.UNKNOWN,
                    value=value,
                    formula=formula,
                    sheet_name=ws.title,
                    col=ox_cell.column,
                    row=ox_cell.row,
                    fill_color=_fill_color(ox_cell),
                ))

        sheets.append(Sheet(
            id=f"sheet.{ws.title.lower().replace(' ', '_')}",
            name=ws.title,
            cells=cells,
        ))

    manifest = Manifest(source_file="", sheet_count=len(sheets))
    return Workbook(sheets=sheets, manifest=manifest)
