import io, json, zipfile
from osheet.parser import parse_xlsx
from osheet.analyzer import run_all
from osheet.emitter.osheet import to_osheet_bytes
from osheet.emitter.xlsx import to_xlsx_bytes

def _analyzed(fixture):
    wb = parse_xlsx(fixture)
    return run_all(wb)

def test_osheet_is_zip(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    assert zipfile.is_zipfile(io.BytesIO(data))

def test_osheet_contains_required_files(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = zf.namelist()
    assert "workbook.json" in names
    assert "sheets.json" in names
    assert "formula_graph.json" in names

def test_osheet_workbook_json_valid(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        obj = json.loads(zf.read("workbook.json"))
    assert "manifest" in obj
    assert obj["manifest"]["sheet_count"] >= 1

def test_xlsx_output_is_valid(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_xlsx_bytes(wb)
    import openpyxl
    loaded = openpyxl.load_workbook(io.BytesIO(data))
    assert len(loaded.sheetnames) >= 1
