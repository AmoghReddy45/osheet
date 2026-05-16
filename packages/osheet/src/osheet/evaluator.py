from __future__ import annotations

import re
import sys
from typing import Any

import numpy as np
import formulas
from openpyxl.utils import get_column_letter, column_index_from_string

from osheet.models import Workbook
from osheet.analyzer.graph import _parse_refs


# Single-cell ref (optionally sheet-qualified). Sheet may be quoted or bare.
_SINGLE_CELL_RE = re.compile(
    r"^\s*(?:(?:'([^']+)'|([A-Za-z_][\w ]*))!)?\s*\$?([A-Z]+)\$?(\d+)\s*$"
)


def _fkey(sheet_name: str, row: int, col: int) -> str:
    """Normalized cell key: SHEETNAME!COLROW (uppercase, no $)."""
    return f"{sheet_name.upper()}!{get_column_letter(col)}{row}"


def _normalize_input_key(input_key: str) -> str:
    """Normalize formulas-library input key to match our value_map keys.

    formulas lib outputs: 'FIXED ASSETS'!D15 (with quotes for sheet names with spaces)
    Our _fkey() outputs:  FIXED ASSETS!D15   (no quotes)
    """
    if input_key.startswith("'"):
        # Find the closing single quote
        bang = input_key.find("'", 1)
        if bang != -1 and input_key[bang + 1 : bang + 2] == "!":
            sheet = input_key[1:bang]
            rest = input_key[bang + 2 :]
            return f"{sheet}!{rest}"
    return input_key


def _expand_range(range_key: str, default_sheet: str, value_map: dict[str, Any]) -> np.ndarray:
    """Turn 'B2:B13' or 'SHEET!B2:SHEET!B13' into a numpy row array for SUM etc."""
    # Strip quotes from sheet name portion if present (e.g. 'FIXED ASSETS'!B2:B13)
    range_key = _normalize_input_key(range_key)
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
        return 0.0  # Excel treats empty/missing cells as 0 in arithmetic
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


def _split_top_level_commas(body: str) -> list[str]:
    """Split a function body at top-level commas, respecting parens and quotes."""
    parts: list[str] = []
    depth = 0
    in_quote = False
    start = 0
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == '"':
            in_quote = not in_quote
        elif not in_quote:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "," and depth == 0:
                parts.append(body[start:i])
                start = i + 1
        i += 1
    parts.append(body[start:])
    return parts


def _find_matching_paren(text: str, open_idx: int) -> int:
    """Given index of '(' in text, return index of matching ')'. -1 if not found.
    Ignores parens inside double-quoted string literals."""
    depth = 0
    in_quote = False
    i = open_idx
    while i < len(text):
        ch = text[i]
        if ch == '"':
            in_quote = not in_quote
        elif not in_quote:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return i
        i += 1
    return -1


def _find_innermost_offset(formula: str) -> tuple[int, int] | None:
    """Find an OFFSET(...) call whose body contains no nested OFFSET(.
    Returns (start_of_OFFSET, index_of_closing_paren) or None."""
    upper = formula.upper()
    search_from = 0
    while True:
        idx = upper.find("OFFSET(", search_from)
        if idx == -1:
            return None
        # Must be at start of token (prev char non-alnum/underscore) to avoid matching e.g. MYOFFSET(
        if idx > 0 and (formula[idx - 1].isalnum() or formula[idx - 1] == "_"):
            search_from = idx + 1
            continue
        open_paren = idx + len("OFFSET")  # points at '('
        close_paren = _find_matching_paren(formula, open_paren)
        if close_paren == -1:
            return None
        body = formula[open_paren + 1 : close_paren]
        if "OFFSET(" not in body.upper():
            return (idx, close_paren)
        # Has nested OFFSET; recurse into body
        inner = _find_innermost_offset(body)
        if inner is None:
            # Shouldn't happen, but be safe
            return (idx, close_paren)
        b_start, b_end = inner
        return (open_paren + 1 + b_start, open_paren + 1 + b_end)


def _eval_subexpr(expr: str, default_sheet: str, value_map: dict[str, Any], parser: Any) -> int:
    func = parser.ast(f"={expr}")[1].compile()
    kwargs: dict[str, Any] = {}
    sheet_prefix = default_sheet.upper() + "!"
    for input_key in func.inputs:
        normalized = _normalize_input_key(input_key)
        if ":" in normalized:
            kwargs[input_key] = _expand_range(normalized, default_sheet.upper(), value_map)
        else:
            lookup = normalized if "!" in normalized else sheet_prefix + normalized
            v = value_map.get(lookup)
            kwargs[input_key] = _to_float(v) if isinstance(v, (int, float, type(None))) else v
    result = _scalar(func(**kwargs))
    return int(result)


def _eval_subexpr_scalar(
    expr: str,
    default_sheet: str,
    value_map: dict[str, Any],
    parser: Any,
    named_tables: dict[str, Any] | None,
    cell_row: int,
) -> Any:
    """Evaluate a sub-expression, returning the raw scalar.

    Unlike ``_eval_subexpr`` this does not cast to int and additionally
    resolves structured table references and OFFSET calls in ``expr``
    before parsing. Any failure returns ``None``.
    """
    try:
        text = expr
        if named_tables and "[" in text:
            text = _resolve_structured_refs(text, cell_row, named_tables)
        if "OFFSET" in text.upper():
            text = _resolve_offset_in_formula(text, default_sheet, value_map, parser)
        if not text.startswith("="):
            text = "=" + text
        func = parser.ast(text)[1].compile()
        kwargs: dict[str, Any] = {}
        sheet_prefix = default_sheet.upper() + "!"
        for input_key in func.inputs:
            normalized = _normalize_input_key(input_key)
            if ":" in normalized:
                kwargs[input_key] = _expand_range(normalized, default_sheet.upper(), value_map)
            else:
                lookup = normalized if "!" in normalized else sheet_prefix + normalized
                v = value_map.get(lookup)
                kwargs[input_key] = _to_float(v) if isinstance(v, (int, float, type(None))) else v
        return _scalar(func(**kwargs))
    except Exception:
        return None


def _try_if_short_circuit(
    formula: str,
    default_sheet: str,
    value_map: dict[str, Any],
    parser: Any,
    named_tables: dict[str, Any] | None,
    cell_row: int,
) -> tuple[bool, Any]:
    """Short-circuit IF(cond, true, false) when an unresolvable OFFSET sits
    in the inactive branch.

    Returns ``(handled, value)``. ``handled=True`` means the result is
    authoritative and the caller should return ``value`` directly.
    ``handled=False`` means the caller should fall back to normal evaluation.
    """
    body = formula.lstrip()
    if body.startswith("="):
        body = body[1:].lstrip()
    if not body.upper().startswith("IF("):
        return (False, None)
    open_paren = body.upper().find("IF(") + 2  # index of '('
    close = _find_matching_paren(body, open_paren)
    if close == -1 or close != len(body.rstrip()) - 1:
        return (False, None)
    inner = body[open_paren + 1 : close]
    parts = _split_top_level_commas(inner)
    if len(parts) != 3:
        return (False, None)
    cond_expr, true_expr, false_expr = (p.strip() for p in parts)
    cond_val = _eval_subexpr_scalar(
        cond_expr, default_sheet, value_map, parser, named_tables, cell_row
    )
    if cond_val is None:
        return (False, None)
    if cond_val == 0 or cond_val is False or cond_val == "":
        is_true = False
    else:
        is_true = bool(cond_val)
    active = true_expr if is_true else false_expr
    active = active.strip()
    # Recurse into nested IF
    if active.upper().startswith("IF("):
        recur_close = _find_matching_paren(active, 2)  # '(' is at index 2
        if recur_close == len(active) - 1:
            handled, val = _try_if_short_circuit(
                "=" + active, default_sheet, value_map, parser, named_tables, cell_row
            )
            if handled:
                return (True, val)
    return (
        True,
        _eval_subexpr_scalar(
            active, default_sheet, value_map, parser, named_tables, cell_row
        ),
    )


def _parse_single_cell_ref(ref: str) -> tuple[str | None, str, int]:
    """Parse 'A1' / '$A$1' / 'Sheet!A1' / ''Sheet Name'!A1' into (sheet_or_None, col_letters, row).
    Raises ValueError if not a single-cell reference."""
    m = _SINGLE_CELL_RE.match(ref)
    if not m:
        raise ValueError(f"not a single cell ref: {ref!r}")
    quoted_sheet, bare_sheet, col_letters, row_str = m.groups()
    sheet = quoted_sheet or bare_sheet
    return sheet, col_letters.upper(), int(row_str)


# Matches `<sheet>!<cell>:<sheet>!<cell>` where sheets are identical.
# Captures both quoted ('Sheet Name') and bare (Sheet1) forms.
_REDUNDANT_SHEET_RE = re.compile(
    r"(?:'([^']+)'|([A-Za-z_][\w ]*))!(\$?[A-Z]+\$?\d+)"
    r":"
    r"(?:'([^']+)'|([A-Za-z_][\w ]*))!(\$?[A-Z]+\$?\d+)"
)


def _collapse_redundant_sheet_in_range(text: str) -> str:
    def repl(m: re.Match) -> str:
        left_sheet = m.group(1) or m.group(2)
        right_sheet = m.group(4) or m.group(5)
        if left_sheet.upper() == right_sheet.upper():
            left_prefix = m.group(0).split(":", 1)[0]
            return f"{left_prefix}:{m.group(6)}"
        return m.group(0)
    return _REDUNDANT_SHEET_RE.sub(repl, text)


def _resolve_offset_in_formula(
    formula: str,
    default_sheet: str,
    value_map: dict[str, Any],
    parser: Any,
) -> str:
    """Rewrite OFFSET(ref, rows, cols, [height], [width]) → concrete cell or range
    in the formula text. On any failure, returns the original formula unchanged."""
    try:
        current = formula
        # Bound iterations to avoid pathological loops
        for _ in range(64):
            upper = current.upper()
            if "OFFSET(" not in upper:
                break
            located = _find_innermost_offset(current)
            if located is None:
                break
            offset_start, close_paren = located
            open_paren = offset_start + len("OFFSET")
            body = current[open_paren + 1 : close_paren]
            args = _split_top_level_commas(body)
            if len(args) < 3:
                return formula  # malformed; bail
            ref_arg = args[0].strip()
            rows_int = _eval_subexpr(args[1].strip(), default_sheet, value_map, parser)
            cols_int = _eval_subexpr(args[2].strip(), default_sheet, value_map, parser)
            height = (
                _eval_subexpr(args[3].strip(), default_sheet, value_map, parser)
                if len(args) >= 4 and args[3].strip()
                else 1
            )
            width = (
                _eval_subexpr(args[4].strip(), default_sheet, value_map, parser)
                if len(args) >= 5 and args[4].strip()
                else 1
            )

            sheet_name, col_letters, base_row = _parse_single_cell_ref(ref_arg)
            base_col = column_index_from_string(col_letters)
            new_col = base_col + cols_int
            new_row = base_row + rows_int
            if new_col < 1 or new_row < 1:
                raise ValueError("OFFSET out of range")
            top_addr = f"{get_column_letter(new_col)}{new_row}"
            if height > 1 or width > 1:
                bottom_col = new_col + max(width, 1) - 1
                bottom_row = new_row + max(height, 1) - 1
                bottom_addr = f"{get_column_letter(bottom_col)}{bottom_row}"
                new_addr = f"{top_addr}:{bottom_addr}"
            else:
                new_addr = top_addr
            if sheet_name:
                prefix = (
                    f"'{sheet_name}'!" if " " in sheet_name else f"{sheet_name}!"
                )
                new_addr = prefix + new_addr

            current = current[:offset_start] + new_addr + current[close_paren + 1 :]

        # Normalize `Sheet!X:Sheet!Y` → `Sheet!X:Y` (formulas lib treats the
        # duplicated-sheet form as two scalars, not a range).
        return _collapse_redundant_sheet_in_range(current)
    except Exception:
        return formula


# Structured-reference pattern: TableName[[#This Row],[ColName]]
_STRUCTURED_REF_RE = re.compile(
    r"([A-Za-z_][\w]*)"             # table name
    r"\[\["                          # [[
    r"#This Row\],\["                # #This Row],[
    r"([^\]]+)"                      # column name (no ])
    r"\]\]"                          # ]]
)


def _resolve_structured_refs(
    formula: str,
    cell_row: int,
    named_tables: dict[str, Any],
) -> str:
    """Rewrite TableName[[#This Row],[ColName]] -> 'Sheet'!COLROW.
    Returns formula unchanged if no rewrite possible."""
    if not named_tables or "[" not in formula:
        return formula

    def _sub(m: re.Match) -> str:
        tname, col_name = m.group(1), m.group(2)
        tbl = named_tables.get(tname)
        if tbl is None or col_name not in tbl.columns:
            return m.group(0)  # leave unchanged -> formula will fail with None
        col = tbl.columns[col_name]
        addr = f"{get_column_letter(col)}{cell_row}"
        sheet = tbl.sheet_name
        prefix = f"'{sheet}'!" if " " in sheet or "'" in sheet else f"{sheet}!"
        return f"{prefix}{addr}"

    return _STRUCTURED_REF_RE.sub(_sub, formula)


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
                if sid:
                    dep_ids.add(sid)
            deps[cell.stable_id] = dep_ids
    return deps


def _tarjan_sccs(deps: dict[str, set[str]]) -> list[list[str]]:
    """Return list of SCCs using Tarjan's algorithm.

    Iteration order of ``all_nodes`` (a set) is non-deterministic in general,
    but because Tarjan's DFS recurses into dependencies before finishing a node,
    dependency SCCs are emitted (appended) before the SCCs that depend on them.
    The result is therefore in sources-first topological order for this
    implementation — suitable for direct use in evaluate_patch without reversal.

    SCCs with len > 1, or len==1 with a self-loop, are circular.
    """
    all_nodes: set[str] = set(deps.keys())
    for s in deps.values():
        all_nodes |= s

    index_counter = [0]
    stack: list[str] = []
    lowlink: dict[str, int] = {}
    index: dict[str, int] = {}
    on_stack: dict[str, bool] = {}
    sccs: list[list[str]] = []

    def strongconnect(v: str) -> None:
        index[v] = lowlink[v] = index_counter[0]
        index_counter[0] += 1
        stack.append(v)
        on_stack[v] = True
        for w in deps.get(v, set()):
            if w not in index:
                strongconnect(w)
                lowlink[v] = min(lowlink[v], lowlink[w])
            elif on_stack.get(w, False):
                lowlink[v] = min(lowlink[v], index[w])
        if lowlink[v] == index[v]:
            scc: list[str] = []
            while True:
                w = stack.pop()
                on_stack[w] = False
                scc.append(w)
                if w == v:
                    break
            sccs.append(scc)

    old_limit = sys.getrecursionlimit()
    sys.setrecursionlimit(max(old_limit, len(all_nodes) * 3 + 1000))
    try:
        for v in all_nodes:
            if v not in index:
                strongconnect(v)
    finally:
        sys.setrecursionlimit(old_limit)
    return sccs


def _eval_one_cell(
    sheet_name: str,
    cell: Any,
    value_map: dict[str, Any],
    id_to_key: dict[str, str],
    parser: Any,
    named_tables: dict[str, Any] | None = None,
) -> Any:
    """Evaluate a single formula cell. Returns scalar or None on error."""
    try:
        formula_text = cell.formula
        if named_tables and "[" in formula_text:
            formula_text = _resolve_structured_refs(formula_text, cell.row, named_tables)
        if "OFFSET" in formula_text.upper():
            formula_text = _resolve_offset_in_formula(
                formula_text, sheet_name.upper(), value_map, parser
            )
            # If the rewrite couldn't resolve every OFFSET (e.g. the shifted
            # column/row would fall out of range) AND the outer wrapper is an
            # IF, short-circuit the IF and skip the dead branch. Excel never
            # evaluates that branch, and the unresolved OFFSET would otherwise
            # cause parse/dispatch failure.
            if "OFFSET(" in formula_text.upper():
                stripped = formula_text.lstrip("=").lstrip()
                if stripped.upper().startswith("IF("):
                    handled, val = _try_if_short_circuit(
                        formula_text,
                        sheet_name.upper(),
                        value_map,
                        parser,
                        named_tables,
                        cell.row,
                    )
                    if handled:
                        return val
        func = parser.ast(formula_text)[1].compile()
        kwargs: dict[str, Any] = {}
        sheet_prefix = sheet_name.upper() + "!"
        for input_key in func.inputs:
            normalized = _normalize_input_key(input_key)
            if ":" in normalized:
                kwargs[input_key] = _expand_range(normalized, sheet_name.upper(), value_map)
            else:
                lookup_key = normalized if "!" in normalized else sheet_prefix + normalized
                v = value_map.get(lookup_key)
                kwargs[input_key] = _to_float(v) if isinstance(v, (int, float, type(None))) else v
        return _scalar(func(**kwargs))
    except Exception:
        return None


def _gauss_seidel(
    scc_cells: list[tuple[str, Any]],
    value_map: dict[str, Any],
    id_to_key: dict[str, str],
    parser: Any,
    max_iterations: int = 50,
    tolerance: float = 1e-6,
    named_tables: dict[str, Any] | None = None,
) -> None:
    """Iterate circular SCC in-place (Gauss-Seidel) until convergence.
    Modifies value_map in place. Silent on non-convergence."""
    # Seed None values with 0 (matches Excel behaviour)
    for _, cell in scc_cells:
        k = id_to_key[cell.stable_id]
        if value_map.get(k) is None:
            value_map[k] = 0.0

    prev_signs: dict[str, int] = {}
    omega = 1.0  # relaxation factor; reduced to 0.5 on detected oscillation

    for _iteration in range(max_iterations):
        max_change = 0.0
        for sheet_name, cell in scc_cells:
            k = id_to_key[cell.stable_id]
            old_val = value_map.get(k)
            new_val = _eval_one_cell(sheet_name, cell, value_map, id_to_key, parser, named_tables)
            if new_val is None:
                continue
            # Apply under-relaxation if oscillation was detected
            if omega < 1.0 and isinstance(old_val, (int, float)):
                new_val = (1.0 - omega) * float(old_val) + omega * float(new_val)
            value_map[k] = new_val  # in-place update (Gauss-Seidel)
            if isinstance(old_val, (int, float)) and isinstance(new_val, (int, float)):
                change = abs(float(new_val) - float(old_val))
                max_change = max(max_change, change)
                # Oscillation: sign of delta flips -> engage damping
                sign = 1 if new_val > old_val else (-1 if new_val < old_val else 0)
                prev = prev_signs.get(cell.stable_id, 0)
                if prev != 0 and sign != 0 and sign != prev:
                    omega = 0.5
                if sign != 0:
                    prev_signs[cell.stable_id] = sign

        if max_change < tolerance:
            break  # converged


def evaluate_patch(
    patches: dict[str, Any],
    workbook: Workbook,
) -> dict[str, Any]:
    """
    Apply {stable_id: new_value} patches and re-evaluate all formula cells.
    Returns {stable_id: computed_value} for every cell in the workbook.
    Uses SCC-based topological evaluation with Gauss-Seidel iteration for
    circular references.
    """
    parser = formulas.Parser()
    named_tables = getattr(workbook, "named_tables", {}) or {}

    # Build value map and bidirectional key mappings
    value_map: dict[str, Any] = {}
    id_to_key: dict[str, str] = {}
    key_to_id: dict[str, str] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            k = _fkey(sheet.name, cell.row, cell.col)
            value_map[k] = cell.value
            id_to_key[cell.stable_id] = k
            key_to_id[k] = cell.stable_id

    # Apply patches
    for stable_id, new_value in patches.items():
        k = id_to_key.get(stable_id)
        if k:
            value_map[k] = new_value

    # Build dependency graph and cell lookup
    dep_graph = _build_dep_graph(workbook)
    cell_lookup: dict[str, tuple[str, Any]] = {}  # stable_id -> (sheet_name, cell)
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if cell.formula:
                cell_lookup[cell.stable_id] = (sheet.name, cell)

    # _tarjan_sccs emits SCCs in sources-first topological order (see its
    # docstring). Dependency SCCs are appended before the SCCs that depend on
    # them, so no reversal is needed here.
    sccs = _tarjan_sccs(dep_graph)
    sccs_topo = sccs

    # Evaluate each SCC in topological order
    for scc in sccs_topo:
        formula_nodes = [sid for sid in scc if sid in cell_lookup]
        if not formula_nodes:
            continue  # SCC contains only non-formula (assumption) cells

        is_circular = len(scc) > 1 or scc[0] in dep_graph.get(scc[0], set())

        if is_circular:
            scc_cells = [cell_lookup[sid] for sid in formula_nodes]
            _gauss_seidel(scc_cells, value_map, id_to_key, parser, named_tables=named_tables)
        else:
            # Non-circular singleton: evaluate once
            sid = formula_nodes[0]
            sheet_name, cell = cell_lookup[sid]
            result = _eval_one_cell(sheet_name, cell, value_map, id_to_key, parser, named_tables)
            value_map[id_to_key[sid]] = result

    return {
        key_to_id[k]: _scalar(v)
        for k, v in value_map.items()
        if k in key_to_id
    }
