from __future__ import annotations
import re
from openpyxl.utils import column_index_from_string
from osheet.models import Cell, Workbook

_RANGE_RE = re.compile(
    r"(?:(?<![A-Za-z0-9_])'?([^'!\[\]]+)'?!)?(\$?[A-Z]+)(\$?[0-9]+):(\$?[A-Z]+)(\$?[0-9]+)",
    re.IGNORECASE,
)
_CELL_RE = re.compile(
    r"(?:(?:'([^'!\[\]]+)'|([A-Za-z_][A-Za-z0-9_]*))\!)?(\$?[A-Z]+)(\$?[0-9]+)(?![\w(])",
    re.IGNORECASE,
)


def _col_num(col_str: str) -> int:
    return column_index_from_string(col_str.replace("$", ""))


def _row_num(row_str: str) -> int:
    return int(row_str.replace("$", ""))


def _expand_range(sheet: str, col1: str, row1: str, col2: str, row2: str) -> list[tuple[str, int, int]]:
    c1, c2 = _col_num(col1), _col_num(col2)
    r1, r2 = _row_num(row1), _row_num(row2)
    refs = []
    for c in range(min(c1, c2), max(c1, c2) + 1):
        for r in range(min(r1, r2), max(r1, r2) + 1):
            refs.append((sheet, c, r))
    return refs


def _parse_refs(formula: str, default_sheet: str) -> list[tuple[str, int, int]]:
    refs: list[tuple[str, int, int]] = []
    range_spans = set()
    for m in _RANGE_RE.finditer(formula):
        sheet = m.group(1) or default_sheet
        refs.extend(_expand_range(sheet, m.group(2), m.group(3), m.group(4), m.group(5)))
        range_spans.add((m.start(), m.end()))
    for m in _CELL_RE.finditer(formula):
        if any(rs <= m.start() and m.end() <= re_ for rs, re_ in range_spans):
            continue
        sheet = m.group(1) or m.group(2) or default_sheet
        try:
            col = _col_num(m.group(3))
            row = _row_num(m.group(4))
            refs.append((sheet, col, row))
        except Exception:
            continue
    return refs


def build_formula_graph(workbook: Workbook) -> None:
    """Mutates each Cell.depends_on with stable_ids of its dependencies."""
    addr_to_id: dict[tuple[str, int, int], str] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            addr_to_id[(sheet.name, cell.col, cell.row)] = cell.stable_id

    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if not cell.formula:
                continue
            refs = _parse_refs(cell.formula, sheet.name)
            deps: list[str] = []
            for (sname, col, row) in refs:
                sid = addr_to_id.get((sname, col, row))
                if sid and sid != cell.stable_id:
                    deps.append(sid)
            cell.depends_on = list(dict.fromkeys(deps))
