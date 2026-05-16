from __future__ import annotations
from typing import Any
import osheet
from mcp.server.fastmcp import FastMCP

_mcp = FastMCP(
    "osheet",
    instructions=(
        "osheet: AI-native spreadsheet compiler. "
        "Load any .xlsx file by its absolute path to inspect assumptions, outputs, "
        "formula dependencies, and simulate what-if patches with exact computed values."
    ),
)
_cache: dict[str, osheet.OsheetWorkbook] = {}


def clear_cache() -> None:
    _cache.clear()


def _load(path: str) -> osheet.OsheetWorkbook:
    if path not in _cache:
        with open(path, "rb") as f:
            _cache[path] = osheet.load(f.read())
    return _cache[path]


def get_workbook_summary(path: str) -> str:
    """Load an xlsx and return a structured summary: sheets, tables, assumption count, output count."""
    wb = _load(path)
    m = wb.manifest
    lines = [
        f"Workbook: {m.sheet_count} sheets, {m.table_count} tables",
        f"Assumptions detected: {m.assumption_count}",
        f"Outputs detected: {m.output_count}",
    ]
    for s in wb.sheets:
        lines.append(f"  Sheet '{s.name}': {len(s.cells)} cells, {len(s.tables)} tables")
    if m.warnings:
        lines.append(f"Warnings ({len(m.warnings)}):")
        for w in m.warnings:
            lines.append(f"  [{w.address}] {w.message}")
    return "\n".join(lines)


def get_assumptions(path: str) -> list[dict[str, Any]]:
    """Return all detected assumption input cells with stable IDs and current values."""
    wb = _load(path)
    return [
        {
            "stable_id": c.stable_id,
            "value": c.value,
            "sheet": c.sheet_name,
            "row": c.row,
            "col": c.col,
            "confidence": round(c.confidence, 2),
        }
        for c in wb.assumptions
    ]


def get_outputs(path: str) -> list[dict[str, Any]]:
    """Return all detected output cells (formula cells that nothing else depends on)."""
    wb = _load(path)
    return [
        {
            "stable_id": c.stable_id,
            "value": c.value,
            "formula": c.formula,
            "sheet": c.sheet_name,
            "row": c.row,
        }
        for c in wb.outputs
    ]


def trace_cell(path: str, stable_id: str) -> dict[str, Any]:
    """Trace a cell's upstream dependencies and downstream dependents."""
    wb = _load(path)
    result = osheet.trace(wb, stable_id)
    return {
        "cell_id": result.cell_id,
        "upstream": result.upstream,
        "downstream": result.downstream,
    }


def find_cells(path: str, query: str) -> list[dict[str, Any]]:
    """Fuzzy search for cells by stable ID fragment or text value substring."""
    wb = _load(path)
    return [
        {
            "stable_id": c.stable_id,
            "value": c.value,
            "role": c.role.value,
            "formula": c.formula,
            "sheet": c.sheet_name,
        }
        for c in osheet.find(wb, query)
    ]


def propose_patch_tool(path: str, stable_id: str, new_value: float) -> dict[str, Any]:
    """
    Show exactly what changes if an assumption is modified.
    Returns affected cell IDs and their new computed values via formula evaluation.
    """
    wb = _load(path)
    proposal = osheet.propose_patch(wb, stable_id, new_value)
    return {
        "cell_id": proposal.cell_id,
        "old_value": proposal.old_value,
        "new_value": proposal.new_value,
        "affected_cells": proposal.affected_cells,
        "computed_values": proposal.computed_values,
        "diff": proposal.diff,
    }


# Register all tools with FastMCP
_mcp.add_tool(get_workbook_summary, name="get_workbook_summary",
    description="Load an xlsx file and return a summary of its structure (sheets, tables, assumption count, output count).")
_mcp.add_tool(get_assumptions, name="get_assumptions",
    description="Return all detected assumption inputs with stable IDs and current values.")
_mcp.add_tool(get_outputs, name="get_outputs",
    description="Return all output cells (formula cells that nothing depends on).")
_mcp.add_tool(trace_cell, name="trace_cell",
    description="Trace a cell's upstream dependencies and downstream dependents by stable ID.")
_mcp.add_tool(find_cells, name="find_cells",
    description="Fuzzy search for cells by stable ID or text value.")
_mcp.add_tool(propose_patch_tool, name="propose_patch",
    description="Show what changes if an assumption changes. Returns exact new computed values via formula evaluation.")


def main() -> None:
    _mcp.run()


if __name__ == "__main__":
    main()
