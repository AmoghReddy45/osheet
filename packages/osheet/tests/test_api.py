import io
import osheet
from osheet.models import CellRole

def test_load_returns_workbook(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    assert len(wb.sheets) >= 1

def test_load_runs_all_passes(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    # After all passes, some cells should have non-UNKNOWN roles
    roles = {c.role for s in wb.sheets for c in s.cells}
    assert roles - {CellRole.UNKNOWN} != set()

def test_trace_returns_upstream(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    outputs = wb.outputs
    if outputs:
        result = osheet.trace(wb, outputs[0].stable_id)
        assert hasattr(result, "upstream")
        assert hasattr(result, "downstream")

def test_find_fuzzy(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    results = osheet.find(wb, "revenue")
    assert isinstance(results, list)

def test_export_xlsx(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    data = wb.export_xlsx()
    assert isinstance(data, bytes)
    assert len(data) > 100

def test_export_osheet(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    data = wb.export_osheet()
    import zipfile
    assert zipfile.is_zipfile(io.BytesIO(data))

def test_propose_patch(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    assumptions = wb.assumptions
    if assumptions:
        a = assumptions[0]
        proposal = osheet.propose_patch(wb, a.stable_id, 999)
        assert proposal.cell_id == a.stable_id
        assert proposal.new_value == 999
        assert isinstance(proposal.affected_cells, list)
