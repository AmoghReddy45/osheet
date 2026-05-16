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
    build_formula_graph(workbook)
    classify_roles(workbook)
    assign_stable_ids(workbook)
    workbook.manifest.table_count = sum(len(s.tables) for s in workbook.sheets)
    workbook.manifest.assumption_count = len(workbook.assumptions)
    workbook.manifest.output_count = len(workbook.outputs)
    workbook.manifest.sheet_count = len(workbook.sheets)
    return workbook
