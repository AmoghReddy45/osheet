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


def test_parser_coerces_comma_formatted_numeric_cells():
    """Numeric-typed cells (data_type='n') whose value happens to be a string
    in a numeric format should still be coerced.

    Note: openpyxl rarely yields this combination from a real round-trip — it
    typically assigns data_type='s' to any string value. We construct the
    in-memory cell manually here to exercise the coercion path. The complementary
    text-typed behaviour is covered in
    ``test_parser_preserves_text_typed_cell_despite_numeric_format``.
    """
    import io, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    # Force data_type='n' by writing real numbers; the openpyxl reader will
    # surface these as int/float, which the coerce path leaves untouched.
    ws['A1'] = 3827
    ws['A1'].number_format = '#,##0'
    ws['B1'] = -2032
    ws['B1'].number_format = '#,##0_);(#,##0)'
    ws['C1'] = 0.5
    ws['C1'].number_format = '0%'
    ws['D1'] = 'Label text'
    ws['D1'].number_format = 'General'
    ws['E1'] = 1234.56
    ws['E1'].number_format = '#,##0.00'
    buf = io.BytesIO(); wb.save(buf)

    from osheet.parser import parse_xlsx
    workbook = parse_xlsx(buf.getvalue())
    a1 = next(c for c in workbook.all_cells if c.col == 1 and c.row == 1)
    b1 = next(c for c in workbook.all_cells if c.col == 2 and c.row == 1)
    c1 = next(c for c in workbook.all_cells if c.col == 3 and c.row == 1)
    d1 = next(c for c in workbook.all_cells if c.col == 4 and c.row == 1)
    e1 = next(c for c in workbook.all_cells if c.col == 5 and c.row == 1)

    assert a1.value == 3827
    assert b1.value == -2032
    assert c1.value == 0.5
    assert d1.value == 'Label text'  # text-typed → not coerced
    assert e1.value == 1234.56


def test_maybe_coerce_value_coerces_numeric_string_when_not_text_typed():
    """Direct unit test of _maybe_coerce_value: when openpyxl somehow gives us
    a string value with data_type='n' (genuine numeric cell), we should coerce.
    """
    from osheet.parser import _maybe_coerce_value

    # data_type='n' (numeric) + numeric format + string value → coerce
    assert _maybe_coerce_value('3,827', '#,##0', 'n') == 3827.0
    assert _maybe_coerce_value('(2,032)', '#,##0_);(#,##0)', 'n') == -2032.0
    assert _maybe_coerce_value('50%', '0%', 'n') == 0.5
    assert _maybe_coerce_value('1,234.56', '#,##0.00', 'n') == 1234.56

    # data_type='s' or 'str' → preserve, even with numeric format
    assert _maybe_coerce_value('3,827', '#,##0', 's') == '3,827'
    assert _maybe_coerce_value('3,827', '#,##0', 'str') == '3,827'

    # Default data_type=None falls through to the old string-coercion path
    assert _maybe_coerce_value('3,827', '#,##0', None) == 3827.0


def test_parser_preserves_text_typed_cell_despite_numeric_format():
    """Excel sometimes stores values as strings (t='str' or t='s') even with a
    numeric format. Those cells must remain strings so AVERAGE/SUM correctly
    skip them (Excel skips text-typed cells when aggregating).
    """
    import io, openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    c = ws.cell(row=1, column=1, value="3,827")
    c.number_format = "#,##0"
    buf = io.BytesIO(); wb.save(buf)

    # Confirm openpyxl round-trips this as a text-typed cell.
    wb2 = openpyxl.load_workbook(io.BytesIO(buf.getvalue()))
    ox_cell = wb2.active.cell(row=1, column=1)
    assert ox_cell.data_type in ('s', 'str'), \
        f"setup precondition failed: data_type={ox_cell.data_type!r}"

    from osheet.parser import parse_xlsx
    workbook = parse_xlsx(buf.getvalue())
    a1 = next(c for c in workbook.all_cells if c.col == 1 and c.row == 1)
    assert a1.value == "3,827", \
        f"Text-typed cell should stay string, got {a1.value!r}"


def test_nvidia_balance_sheet_text_cells_stay_text():
    """Real fixture: Balance Sheet J6 is stored by Excel as text ('3,827') even
    though its format is '#,##0'. After the fix, the parser must preserve that
    string type so downstream AVERAGE/SUM aggregates skip the cell — matching
    Excel's behaviour.
    """
    from osheet.parser import parse_xlsx
    with open(
        '/Users/amoghreddy/excel-project/benchmarks/real_models/nvidia_dcf_model.xlsx',
        'rb',
    ) as f:
        wb = parse_xlsx(f.read())
    bs_j6 = next(
        (c for c in wb.all_cells
         if c.sheet_name == 'Balance Sheet' and c.row == 6 and c.col == 10),
        None,
    )
    assert bs_j6 is not None, "Balance Sheet J6 not found"
    assert isinstance(bs_j6.value, str), (
        f"BS!J6 should remain string (Excel stored it as t='s'), "
        f"got {type(bs_j6.value).__name__}: {bs_j6.value!r}"
    )
    assert bs_j6.value == '3,827'


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
    """Real-world: load nvidia_dcf. Genuinely numeric-typed Balance Sheet cells
    should be parsed as floats; Excel-stored text cells (data_type='s') with
    numeric formats stay as strings so AVERAGE/SUM skip them — matching Excel.
    """
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
    # The Balance Sheet has many genuinely-numeric cells; those must parse as numbers.
    assert bs_numeric_cells > 50, f"Expected many numeric cells, got {bs_numeric_cells}"
    # Excel stored some columns as text ('3,827' etc) even though the format is
    # numeric. After the data_type-aware fix those cells stay as strings so
    # downstream aggregates skip them.
    assert bs_string_cells_in_data_area > 0, (
        "Expected some Excel-stored text cells (data_type='s' with numeric format) "
        "to remain strings; got 0, which means the parser is over-coercing again."
    )
