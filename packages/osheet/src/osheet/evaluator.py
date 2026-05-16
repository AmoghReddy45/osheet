from __future__ import annotations

from typing import Any

import numpy as np
import formulas
from openpyxl.utils import get_column_letter, column_index_from_string

from osheet.models import Workbook
from osheet.analyzer.graph import _parse_refs


def _fkey(sheet_name: str, row: int, col: int) -> str:
    """Normalized cell key: SHEETNAME!COLROW (uppercase, no $)."""
    return f"{sheet_name.upper()}!{get_column_letter(col)}{row}"


def _expand_range(range_key: str, default_sheet: str, value_map: dict[str, Any]) -> np.ndarray:
    """Turn 'B2:B13' or 'SHEET!B2:SHEET!B13' into a numpy row array for SUM etc."""
    if "!" in range_key:
        left, right = range_key.split(":", 1)
        sheet = left.split("!")[0]
        start_cell = left.split("!")[1]
        end_cell = right.split("!")[1] if "!" in right else right
    else:
        sheet = default_sheet
        start_cell, end_cell = range_key.split(":", 1)

    c1 = column_index_from_string("".join(c for c in start_cell if c.isalpha()))
    r1 = int("".join(c for c in start_cell if c.isdigit()))
    c2 = column_index_from_string("".join(c for c in end_cell if c.isalpha()))
    r2 = int("".join(c for c in end_cell if c.isdigit()))

    vals = []
    for row in range(min(r1, r2), max(r1, r2) + 1):
        for col in range(min(c1, c2), max(c1, c2) + 1):
            v = value_map.get(f"{sheet.upper()}!{get_column_letter(col)}{row}", np.nan)
            vals.append(_to_float(v))
    return np.array([vals])


def _to_float(val: Any) -> float:
    if val is None:
        return np.nan
    if isinstance(val, (int, float)):
        return float(val)
    # Try numpy scalar extraction
    try:
        arr = np.array(val)
        return float(arr.item())
    except Exception:
        pass
    if hasattr(val, "__iter__") and not isinstance(val, (str, bytes)):
        try:
            return float(list(list(val)[0])[0])
        except Exception:
            return np.nan
    try:
        return float(val)
    except (TypeError, ValueError):
        return np.nan


def _scalar(val: Any) -> Any:
    """Extract a Python scalar from a formulas Array or numpy array."""
    if val is None:
        return None
    if isinstance(val, (int, float, bool)):
        return val
    # Try numpy scalar extraction first (works for 0-d arrays and Array objects)
    try:
        arr = np.array(val)
        return arr.item()
    except Exception:
        pass
    # Fallback: iterate
    if hasattr(val, "__iter__") and not isinstance(val, (str, bytes)):
        try:
            inner = list(val)
            if inner and hasattr(inner[0], "__iter__"):
                inner = list(inner[0])
            v = inner[0]
            return float(v) if isinstance(v, (int, float, np.floating)) else v
        except (IndexError, TypeError):
            return None
    return val


def _build_dep_graph(
    workbook: Workbook,
) -> dict[str, set[str]]:
    """
    Build {stable_id -> set of stable_ids it depends on} using formula text,
    bypassing the (potentially stale) cell.depends_on field.
    """
    # Position lookup: (sheet_name, col, row) -> stable_id
    pos_to_id: dict[tuple[str, int, int], str] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            pos_to_id[(sheet.name, cell.col, cell.row)] = cell.stable_id

    deps: dict[str, set[str]] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if not cell.formula:
                continue
            refs = _parse_refs(cell.formula, sheet.name)
            dep_ids: set[str] = set()
            for (sname, col, row) in refs:
                sid = pos_to_id.get((sname, col, row))
                if sid and sid != cell.stable_id:
                    dep_ids.add(sid)
            deps[cell.stable_id] = dep_ids
    return deps


def evaluate_patch(
    patches: dict[str, Any],
    workbook: Workbook,
) -> dict[str, Any]:
    """
    Apply {stable_id: new_value} patches and re-evaluate all formula cells.

    Returns {stable_id: computed_value} for every cell in the workbook.
    workbook is not mutated — call evaluate_patch again to re-evaluate from scratch.
    Formula evaluation is in topological order derived from formula references.
    """
    parser = formulas.Parser()

    # Build value map: normalized_key -> current value
    value_map: dict[str, Any] = {}
    id_to_key: dict[str, str] = {}
    key_to_id: dict[str, str] = {}

    for sheet in workbook.sheets:
        for cell in sheet.cells:
            k = _fkey(sheet.name, cell.row, cell.col)
            value_map[k] = cell.value
            id_to_key[cell.stable_id] = k
            key_to_id[k] = cell.stable_id

    # Apply patches (override current values)
    for stable_id, new_value in patches.items():
        k = id_to_key.get(stable_id)
        if k:
            value_map[k] = new_value

    # Build accurate dependency graph from formula text
    dep_graph = _build_dep_graph(workbook)

    # Topological sort: evaluate cells whose dependencies are already resolved
    non_formula_ids = {
        c.stable_id
        for s in workbook.sheets
        for c in s.cells
        if not c.formula
    }
    resolved: set[str] = set(non_formula_ids)

    formula_cells = [
        (s.name, c)
        for s in workbook.sheets
        for c in s.cells
        if c.formula
    ]
    ordered: list[tuple[str, Any]] = []
    remaining = list(formula_cells)

    for _ in range(len(formula_cells) + 1):
        if not remaining:
            break
        next_round = []
        for sheet_name, cell in remaining:
            cell_deps = dep_graph.get(cell.stable_id, set())
            if all(d in resolved for d in cell_deps):
                ordered.append((sheet_name, cell))
                resolved.add(cell.stable_id)
            else:
                next_round.append((sheet_name, cell))
        if len(next_round) == len(remaining):
            # Genuine cycle — mark all cyclic cells as None, skip evaluation
            for sheet_name, cell in next_round:
                value_map[id_to_key[cell.stable_id]] = None
                resolved.add(cell.stable_id)
            break
        remaining = next_round

    # Evaluate each formula cell in dependency order
    for sheet_name, cell in ordered:
        try:
            func = parser.ast(cell.formula)[1].compile()
            kwargs: dict[str, Any] = {}
            sheet_prefix = sheet_name.upper() + "!"
            for input_key in func.inputs:
                if ":" in input_key:
                    kwargs[input_key] = _expand_range(
                        input_key, sheet_name.upper(), value_map
                    )
                else:
                    # Qualify unqualified cell refs with the current sheet
                    lookup_key = (
                        input_key
                        if "!" in input_key
                        else sheet_prefix + input_key
                    )
                    v = value_map.get(lookup_key)
                    kwargs[input_key] = (
                        _to_float(v) if isinstance(v, (int, float, type(None))) else v
                    )
            result = _scalar(func(**kwargs))
            value_map[id_to_key[cell.stable_id]] = result
        except Exception:
            value_map[id_to_key[cell.stable_id]] = None

    return {
        key_to_id[k]: _scalar(v)
        for k, v in value_map.items()
        if k in key_to_id
    }
