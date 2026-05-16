from __future__ import annotations
from osheet.models import Cell, CellRole, Workbook

# RGB-only (last 6 hex chars) colors that indicate an assumption/input cell
_ASSUMPTION_RGB = {
    "FFFF00",  # yellow
    "FFC000",  # orange-yellow
    "FFEB9C",  # light yellow
    "D9E1F2",  # light blue (sometimes used for inputs)
}


def _is_assumption_color(fill_color: str | None) -> bool:
    if not fill_color:
        return False
    # openpyxl returns 8-char ARGB strings like "00FFFF00" or "FFFFFF00"
    # Normalise to just the RGB part (last 6 chars) for comparison
    rgb = fill_color.upper()[-6:]
    return rgb in _ASSUMPTION_RGB


def classify_roles(workbook: Workbook) -> None:
    """Mutates each Cell.role in place using graph in/out-degree + fill heuristics."""
    all_cells = workbook.all_cells
    cell_by_id = {c.stable_id: c for c in all_cells}

    # Count in-degree (how many cells reference this cell)
    in_degree: dict[str, int] = {c.stable_id: 0 for c in all_cells}
    for cell in all_cells:
        for dep_id in cell.depends_on:
            if dep_id in in_degree:
                in_degree[dep_id] += 1

    for cell in all_cells:
        # Labels: text cells with no formula
        if isinstance(cell.value, str) and not cell.formula:
            cell.role = CellRole.LABEL
            continue

        # Assumptions: hardcoded scalar (no formula), referenced by others
        if not cell.formula:
            if in_degree.get(cell.stable_id, 0) > 0:
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.75
            elif _is_assumption_color(cell.fill_color):
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.9
            elif isinstance(cell.value, (int, float)):
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.5
            else:
                cell.role = CellRole.UNKNOWN
            continue

        # Formula cells
        out_degree = len(cell.depends_on)
        dependents_count = in_degree.get(cell.stable_id, 0)

        if dependents_count == 0:
            cell.role = CellRole.OUTPUT
            cell.confidence = 0.85
        elif out_degree > 0:
            cell.role = CellRole.INTERMEDIATE
            cell.confidence = 0.8
        else:
            cell.role = CellRole.UNKNOWN

        # Override: assumption-colored fill on formula cell → probably a special output
        if _is_assumption_color(cell.fill_color):
            cell.role = CellRole.ASSUMPTION
            cell.confidence = 0.7
