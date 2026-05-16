from osheet.parser import parse_xlsx


def test_parse_sheets(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    assert len(raw.sheets) == 1
    assert raw.sheets[0].name == "Revenue"


def test_parse_cells_count(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    assert len(sheet.cells) > 0


def test_parse_formula_cell(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    formula_cells = [c for c in sheet.cells if c.formula]
    assert len(formula_cells) >= 3  # D5, D6, D7 + B10


def test_parse_fill_color(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    b3 = next((c for c in sheet.cells if c.row == 3 and c.col == 2), None)
    assert b3 is not None
    assert b3.fill_color is not None


def test_cross_sheet(cross_sheet_xlsx):
    raw = parse_xlsx(cross_sheet_xlsx)
    assert len(raw.sheets) == 2
    names = [s.name for s in raw.sheets]
    assert "Inputs" in names and "Revenue" in names


def test_formula_value_is_none(simple_xlsx):
    """Formula cells should have value=None (we don't evaluate formulas)."""
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    formula_cells = [c for c in sheet.cells if c.formula]
    for fc in formula_cells:
        assert fc.value is None


def test_scalar_value_preserved(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    # B3 = 0.04 (assumption value)
    b3 = next((c for c in sheet.cells if c.row == 3 and c.col == 2), None)
    assert b3 is not None
    assert b3.value == 0.04
    assert b3.formula is None
