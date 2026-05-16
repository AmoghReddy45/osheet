from .tables import detect_tables
from .types import infer_column_types
from .roles import classify_roles
from .graph import build_formula_graph
from .ids import assign_stable_ids


def run_all(workbook):
    """Run all 5 analyzer passes in sequence, mutating workbook in place."""
    for sheet in workbook.sheets:
        tables = detect_tables(sheet)
        sheet.tables = tables
        infer_column_types(sheet)
    # First pass: populate depends_on so classify_roles can compute in/out-degree.
    # At this point stable_ids are raw (e.g. "Income Statement.C33"), which is fine
    # because classify_roles only needs degree counts, not the IDs themselves.
    build_formula_graph(workbook)
    classify_roles(workbook)
    assign_stable_ids(workbook)
    # Second pass: rewrite depends_on using the final stable_ids so that the BFS
    # in propose_patch (which matches on stable_id) can traverse the graph correctly.
    build_formula_graph(workbook)
    workbook.manifest.table_count = sum(len(s.tables) for s in workbook.sheets)
    workbook.manifest.assumption_count = len(workbook.assumptions)
    workbook.manifest.output_count = len(workbook.outputs)
    workbook.manifest.sheet_count = len(workbook.sheets)
    return workbook
