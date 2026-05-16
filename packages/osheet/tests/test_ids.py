from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables
from osheet.analyzer.ids import assign_stable_ids
import re

def test_ids_are_unique(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    for sheet in raw.sheets:
        sheet.tables = detect_tables(sheet)
    assign_stable_ids(raw)
    all_ids = [c.stable_id for s in raw.sheets for c in s.cells]
    assert len(all_ids) == len(set(all_ids))

def test_ids_are_slugified(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    for sheet in raw.sheets:
        sheet.tables = detect_tables(sheet)
    assign_stable_ids(raw)
    for cell in raw.all_cells:
        assert re.match(r'^[a-z0-9._]+$', cell.stable_id), f"Bad ID: {cell.stable_id}"

def test_ids_stable_across_runs(simple_xlsx):
    raw1 = parse_xlsx(simple_xlsx)
    raw2 = parse_xlsx(simple_xlsx)
    for wb in (raw1, raw2):
        for sheet in wb.sheets:
            sheet.tables = detect_tables(sheet)
        assign_stable_ids(wb)
    ids1 = sorted(c.stable_id for s in raw1.sheets for c in s.cells)
    ids2 = sorted(c.stable_id for s in raw2.sheets for c in s.cells)
    assert ids1 == ids2
