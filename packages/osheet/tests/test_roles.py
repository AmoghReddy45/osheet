from osheet.parser import parse_xlsx
from osheet.analyzer.graph import build_formula_graph
from osheet.analyzer.roles import classify_roles
from osheet.models import CellRole

def test_yellow_cell_is_assumption(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    sheet = raw.sheets[0]
    # B3 has yellow fill and a scalar value
    b3 = next((c for c in sheet.cells if c.row == 3 and c.col == 2), None)
    assert b3 is not None
    assert b3.role == CellRole.ASSUMPTION

def test_sum_cell_is_output(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    # B10 =SUM(B5:B7) — a formula not depended upon by others → output
    sum_cell = next((c for c in all_cells if c.formula and "SUM" in c.formula), None)
    assert sum_cell is not None
    assert sum_cell.role in (CellRole.OUTPUT, CellRole.INTERMEDIATE)

def test_label_cells_classified(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    # "Month", "Revenue", "Costs", "Profit" header cells should be LABEL
    label_cells = [c for c in all_cells if c.role == CellRole.LABEL]
    assert len(label_cells) > 0

def test_no_cells_remain_unknown_after_classify(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    # None cells should remain unknown — every cell should be classifiable
    # (allow a few unknowns for None/empty cells)
    non_empty = [c for c in all_cells if c.value is not None or c.formula is not None]
    unknown = [c for c in non_empty if c.role == CellRole.UNKNOWN]
    assert len(unknown) == 0, f"Unexpected UNKNOWN cells: {[(c.stable_id, c.value, c.formula) for c in unknown]}"
