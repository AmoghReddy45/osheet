from osheet.parser import parse_xlsx
from osheet.analyzer.graph import build_formula_graph


def test_formula_creates_edge(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    formula_cells = [c for s in raw.sheets for c in s.cells if c.formula]
    assert any(len(c.depends_on) > 0 for c in formula_cells)


def test_cross_sheet_dependency(cross_sheet_xlsx):
    raw = parse_xlsx(cross_sheet_xlsx)
    build_formula_graph(raw)
    rev_sheet = next(s for s in raw.sheets if s.name == "Revenue")
    b2 = next((c for c in rev_sheet.cells if c.row == 2 and c.col == 2), None)
    assert b2 is not None
    assert len(b2.depends_on) >= 1


def test_sum_range_expands(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    sum_cell = next((c for c in all_cells if c.formula and "SUM" in (c.formula or "")), None)
    if sum_cell:
        assert len(sum_cell.depends_on) >= 1


def test_no_self_reference(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    for s in raw.sheets:
        for c in s.cells:
            assert c.stable_id not in c.depends_on
