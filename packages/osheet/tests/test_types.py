from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables
from osheet.analyzer.types import infer_column_types
from osheet.models import ColumnDtype


def test_number_column(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    sheet.tables = detect_tables(sheet)
    infer_column_types(sheet)
    for table in sheet.tables:
        for col in table.columns:
            if "Revenue" in col.name:
                assert col.dtype == ColumnDtype.NUMBER


def test_formula_column(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    sheet.tables = detect_tables(sheet)
    infer_column_types(sheet)
    for table in sheet.tables:
        for col in table.columns:
            if "Profit" in col.name:
                assert col.dtype in (ColumnDtype.FORMULA, ColumnDtype.NUMBER, ColumnDtype.UNKNOWN)


def test_unknown_for_empty_column():
    from osheet.models import Column, Table, Sheet
    sheet = Sheet(id="sheet.x", name="X", cells=[],
                  tables=[Table(id="t", sheet_name="X", range_ref="A1:B2",
                                columns=[Column(name="Empty", col_index=1)],
                                header_row=1, first_data_row=2, last_data_row=2)])
    infer_column_types(sheet)
    assert sheet.tables[0].columns[0].dtype == ColumnDtype.UNKNOWN
