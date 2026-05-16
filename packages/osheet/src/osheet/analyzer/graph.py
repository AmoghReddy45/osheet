from __future__ import annotations
import re
from openpyxl.utils import column_index_from_string
from osheet.models import Cell, Workbook

# Strict atom-pair regex: matches `<atom>:<atom>` where each atom is an optionally
# sheet-qualified (quoted or bare) cell reference with optional `$` anchors.
# Anchored at non-identifier boundaries so things like `=SUM(` aren't eaten by
# greedy sheet-name matching.
_RANGE_RE = re.compile(
    r"(?<![A-Za-z0-9_$])"
    r"((?:'[^']+'|[A-Za-z_]\w*)?!?\$?[A-Za-z]+\$?\d+"
    r"\s*:\s*"
    r"(?:'[^']+'|[A-Za-z_]\w*)?!?\$?[A-Za-z]+\$?\d+)"
    r"(?![A-Za-z0-9_$])"
)
_CELL_RE = re.compile(
    r"(?:(?:'([^'!\[\]]+)'|([A-Za-z_][A-Za-z0-9_]*))\!)?(\$?[A-Z]+)(\$?[0-9]+)(?![\w(])",
    re.IGNORECASE,
)

# Single-atom regex used by _parse_range_endpoints to dissect each endpoint.
# If a sheet prefix is present, the `!` is mandatory — otherwise a bare `A10`
# would be parsed as sheet=A, col=Q, row=10 etc. Unquoted sheet names are
# permissive (Excel allows ``&``, ``-``, ``.``, spaces and other chars after
# the first letter). The unquoted branch starts with a letter/underscore so it
# can't begin with a digit and uses ``[^!]*`` so it consumes everything up to
# the mandatory ``!`` delimiter.
_RANGE_ATOM_RE = re.compile(
    r"^(?:(?:'([^']+)'|([A-Za-z_][^!]*))!)?\s*\$?([A-Za-z]+)\$?(\d+)$"
)


def _col_num(col_str: str) -> int:
    return column_index_from_string(col_str.replace("$", ""))


def _row_num(row_str: str) -> int:
    return int(row_str.replace("$", ""))


def _parse_range_endpoints(
    range_str: str, default_sheet: str
) -> tuple[str, int, int, int, int] | None:
    """Parse a range like ``A1:C5`` / ``Sheet!A1:Sheet!C5`` /
    ``TBA!$D$10:'TBA'!I10`` / ``(TBA!D10: TBA!I10)`` into
    ``(sheet, c1, r1, c2, r2)``. Sheet name is returned with its original
    casing (callers uppercase as needed). Returns ``None`` if not parseable.
    """
    s = range_str.strip()
    # Strip wrapping parens (formulas lib emits e.g. "(TBA!D10: TBA!I10)")
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1].strip()
    if ":" not in s:
        return None
    left, _, right = s.partition(":")
    left, right = left.strip(), right.strip()

    lm = _RANGE_ATOM_RE.match(left)
    rm = _RANGE_ATOM_RE.match(right)
    if not lm or not rm:
        return None
    l_sheet = (lm.group(1) or lm.group(2) or default_sheet).strip()
    # Right side sheet is allowed to be missing (e.g. ``Sheet!A1:C5``).
    # If both sides specify a sheet and they disagree, the left wins (Excel
    # range op binds to one sheet — out-of-spec, but be defensive).
    sheet = l_sheet
    c1 = column_index_from_string(lm.group(3).upper())
    r1 = int(lm.group(4))
    c2 = column_index_from_string(rm.group(3).upper())
    r2 = int(rm.group(4))
    return (sheet, c1, r1, c2, r2)


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
    # Scrubbed copy of the formula with each matched range replaced by spaces,
    # so the subsequent _CELL_RE pass doesn't double-count endpoints.
    scrubbed = formula
    for m in _RANGE_RE.finditer(formula):
        endpoints = _parse_range_endpoints(m.group(1), default_sheet)
        if endpoints is None:
            continue
        sheet, c1, r1, c2, r2 = endpoints
        for c in range(min(c1, c2), max(c1, c2) + 1):
            for r in range(min(r1, r2), max(r1, r2) + 1):
                refs.append((sheet, c, r))
        scrubbed = scrubbed[: m.start()] + (" " * (m.end() - m.start())) + scrubbed[m.end() :]
    for m in _CELL_RE.finditer(scrubbed):
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
