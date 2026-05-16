from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables


def test_detects_table_in_simple_model(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    assert len(tables) >= 1


def test_table_has_range(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    t = tables[0]
    assert ":" in t.range_ref


def test_table_header_row(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    t = tables[0]
    assert t.header_row >= 1


def test_table_has_columns(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    t = tables[0]
    assert len(t.columns) >= 2


def test_empty_sheet_returns_no_tables():
    from osheet.models import Sheet
    sheet = Sheet(id="sheet.empty", name="Empty", cells=[])
    assert detect_tables(sheet) == []
