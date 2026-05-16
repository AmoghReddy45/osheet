from osheet.models import Cell, CellRole, Sheet, Table, Workbook, Manifest


def test_cell_defaults():
    c = Cell(stable_id="sheet.A1", role=CellRole.UNKNOWN, value=42)
    assert c.formula is None
    assert c.depends_on == []
    assert c.confidence == 0.0


def test_workbook_flat_indexes():
    assumption = Cell(stable_id="s.A1", role=CellRole.ASSUMPTION, value=0.04)
    output = Cell(stable_id="s.B1", role=CellRole.OUTPUT, value=100.0)
    sheet = Sheet(id="sheet.s", name="S", tables=[], cells=[assumption, output])
    wb = Workbook(sheets=[sheet], manifest=Manifest())
    assert len(wb.assumptions) == 1
    assert len(wb.outputs) == 1


def test_get_cell():
    cell = Cell(stable_id="revenue.b1", role=CellRole.ASSUMPTION, value=100)
    sheet = Sheet(id="sheet.revenue", name="Revenue", cells=[cell])
    wb = Workbook(sheets=[sheet])
    found = wb.get_cell("revenue.b1")
    assert found is not None
    assert found.value == 100


def test_all_cells_across_sheets():
    c1 = Cell(stable_id="s1.a1", role=CellRole.UNKNOWN, value=1)
    c2 = Cell(stable_id="s2.a1", role=CellRole.UNKNOWN, value=2)
    wb = Workbook(sheets=[
        Sheet(id="sheet.s1", name="S1", cells=[c1]),
        Sheet(id="sheet.s2", name="S2", cells=[c2]),
    ])
    assert len(wb.all_cells) == 2


def test_cell_address_property():
    c = Cell(stable_id="x", sheet_name="Revenue", col=2, row=5)
    assert c.address == "Revenue!B5"
