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


def test_parser_coerces_comma_formatted_text():
    import io, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = '3,827'
    ws['A1'].number_format = '#,##0'
    ws['B1'] = '(2,032)'
    ws['B1'].number_format = '#,##0_);(#,##0)'
    ws['C1'] = '50%'
    ws['C1'].number_format = '0%'
    ws['D1'] = 'Label text'
    ws['D1'].number_format = 'General'
    ws['E1'] = '1,234.56'
    ws['E1'].number_format = '#,##0.00'
    buf = io.BytesIO(); wb.save(buf)

    from osheet.parser import parse_xlsx
    workbook = parse_xlsx(buf.getvalue())
    a1 = next(c for c in workbook.all_cells if c.col == 1 and c.row == 1)
    b1 = next(c for c in workbook.all_cells if c.col == 2 and c.row == 1)
    c1 = next(c for c in workbook.all_cells if c.col == 3 and c.row == 1)
    d1 = next(c for c in workbook.all_cells if c.col == 4 and c.row == 1)
    e1 = next(c for c in workbook.all_cells if c.col == 5 and c.row == 1)

    assert a1.value == 3827.0
    assert b1.value == -2032.0
    assert c1.value == 0.5
    assert d1.value == 'Label text'  # text format → not coerced
    assert e1.value == 1234.56


def test_parser_preserves_text_with_general_format():
    """Strings with General format should remain strings even if they look numeric."""
    import io, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = '123'
    ws['A1'].number_format = 'General'
    buf = io.BytesIO(); wb.save(buf)

    from osheet.parser import parse_xlsx
    workbook = parse_xlsx(buf.getvalue())
    a1 = next(c for c in workbook.all_cells if c.col == 1 and c.row == 1)
    # General format: numbers are stored as numbers by openpyxl, strings as strings
    # If openpyxl gave us a string here, it's because the underlying xml stored it that way
    # In that edge case, we leave it alone
    assert a1.value == '123'


def test_parser_handles_unparseable_text_in_numeric_format():
    """A string that doesn't parse as a number in a numeric format → keep as string (don't crash)."""
    import io, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = 'N/A'
    ws['A1'].number_format = '#,##0'
    buf = io.BytesIO(); wb.save(buf)

    from osheet.parser import parse_xlsx
    workbook = parse_xlsx(buf.getvalue())
    a1 = next(c for c in workbook.all_cells if c.col == 1 and c.row == 1)
    assert a1.value == 'N/A'


def test_nvidia_balance_sheet_values_become_numeric():
    """Real-world: load nvidia_dcf and check Balance Sheet values are floats not strings."""
    from osheet.parser import parse_xlsx
    with open('/Users/amoghreddy/excel-project/benchmarks/real_models/nvidia_dcf_model.xlsx', 'rb') as f:
        wb = parse_xlsx(f.read())
    bs_numeric_cells = 0
    bs_string_cells_in_data_area = 0
    for c in wb.all_cells:
        if c.sheet_name != 'Balance Sheet': continue
        if c.formula: continue
        if c.row < 2: continue  # skip header
        if isinstance(c.value, (int, float)):
            bs_numeric_cells += 1
        elif isinstance(c.value, str) and not c.value.replace(' ', '').replace('-', '').replace('(', '').replace(')', '').replace(',', '').replace('.', '').replace('$', '').isalpha():
            # Looks numeric but is a string
            bs_string_cells_in_data_area += 1
    # After fix, there should be many numeric cells and few or no numeric-looking strings
    assert bs_numeric_cells > 50, f"Expected many numeric cells, got {bs_numeric_cells}"
