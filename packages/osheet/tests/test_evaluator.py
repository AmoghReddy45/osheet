# packages/osheet/tests/test_evaluator.py
import io
import openpyxl
import pytest
from osheet.parser import parse_xlsx
from osheet.analyzer import run_all
from osheet.evaluator import evaluate_patch


@pytest.fixture
def two_sheet_bytes() -> bytes:
    """Config (assumptions) + Revenue (formula cells)."""
    wb = openpyxl.Workbook()
    ws_c = wb.active
    ws_c.title = "Config"
    ws_c["A1"] = "growth_rate"
    ws_c["B1"] = 0.10          # assumption — 10% growth
    ws_c["A2"] = "base_arr"
    ws_c["B2"] = 100_000       # assumption — 100k base

    ws_r = wb.create_sheet("Revenue")
    ws_r["A1"] = "Year 1"
    ws_r["B1"] = "=Config!B2*(1+Config!B1)"   # = 110,000
    ws_r["A2"] = "Year 2"
    ws_r["B2"] = "=B1*(1+Config!B1)"           # = 121,000
    ws_r["A3"] = "Total"
    ws_r["B3"] = "=SUM(B1:B2)"                 # = 231,000

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_evaluate_baseline_values(two_sheet_bytes):
    workbook = parse_xlsx(two_sheet_bytes)
    run_all(workbook)
    # No patch — should return current values
    result = evaluate_patch({}, workbook)
    year1 = next(c for c in workbook.all_cells if c.formula and "Config!B2" in c.formula and "Config!B1" in c.formula)
    assert abs(result[year1.stable_id] - 110_000) < 1


def test_evaluate_patch_growth_rate(two_sheet_bytes):
    workbook = parse_xlsx(two_sheet_bytes)
    run_all(workbook)

    growth = next(c for c in workbook.all_cells if c.value == 0.10)
    result = evaluate_patch({growth.stable_id: 0.20}, workbook)

    # Year 1 = 100000 * 1.20 = 120,000
    year1 = next(c for c in workbook.all_cells if c.formula and "Config!B2" in c.formula)
    assert abs(result[year1.stable_id] - 120_000) < 1


def test_evaluate_patch_base_arr(two_sheet_bytes):
    workbook = parse_xlsx(two_sheet_bytes)
    run_all(workbook)

    base = next(c for c in workbook.all_cells if c.value == 100_000)
    result = evaluate_patch({base.stable_id: 200_000}, workbook)

    year1 = next(c for c in workbook.all_cells if c.formula and "Config!B2" in c.formula)
    assert abs(result[year1.stable_id] - 220_000) < 1


def test_evaluate_sum_propagates(two_sheet_bytes):
    workbook = parse_xlsx(two_sheet_bytes)
    run_all(workbook)

    growth = next(c for c in workbook.all_cells if c.value == 0.10)
    result = evaluate_patch({growth.stable_id: 0.20}, workbook)

    # Total = Year1 + Year2 = 120000 + 144000 = 264000
    total = next(c for c in workbook.all_cells if c.formula and "SUM" in c.formula)
    assert abs(result[total.stable_id] - 264_000) < 1


def test_evaluate_bad_formula_returns_none(two_sheet_bytes):
    """Cells with unparseable formulas should return None, not the stale value."""
    workbook = parse_xlsx(two_sheet_bytes)
    run_all(workbook)
    # Find a real formula cell and inject a bad formula
    formula_cell = next(c for c in workbook.all_cells if c.formula)
    original_formula = formula_cell.formula
    formula_cell.formula = "=)BROKEN("
    formula_cell.value = 99999  # stale value before bad formula
    result = evaluate_patch({}, workbook)
    # Restore
    formula_cell.formula = original_formula
    # Should return None for failed formula, not the stale 99999
    assert result[formula_cell.stable_id] is None


@pytest.fixture
def self_ref_bytes() -> bytes:
    """A1 = A1 (self-reference, seeds to 0, converges to 0)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "=A1"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


@pytest.fixture
def two_cell_cycle_bytes() -> bytes:
    """
    A1 = 1000 (hardcoded)
    B1 = (A1 + C1) / 2     depends on C1
    C1 = B1 * 0.1           depends on B1  --> circular B1<->C1
    Fixed point: B1 = 1000/1.9 ≈ 526.3158, C1 = B1*0.1 ≈ 52.6316
    """
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = 1000
    ws["B1"] = "=(A1+C1)/2"
    ws["C1"] = "=B1*0.1"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


@pytest.fixture
def interest_chain_bytes() -> bytes:
    """
    Classic 3-statement interest circularity (simplified):
    A1 = 10_000   (EBITDA, hardcoded)
    A2 = 0.05     (interest rate, hardcoded)
    A3 = A1 - A4  (EBT = EBITDA - interest expense; depends on A4 circular)
    A4 = A5 * A2  (interest expense = avg_debt * rate; depends on A5 circular)
    A5 = A3 * 2   (avg_debt = EBT * 2, simplified; depends on A3 circular)
    Fixed point:
      A4 = A5*0.05 = A3*2*0.05 = A3*0.1
      A3 = A1 - A4 = 10000 - A3*0.1
      1.1*A3 = 10000 -> A3 = 9090.909, A4 = 909.09, A5 = 18181.8
    """
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = 10_000
    ws["A2"] = 0.05
    ws["A3"] = "=A1-A4"
    ws["A4"] = "=A5*A2"
    ws["A5"] = "=A3*2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


@pytest.fixture
def diverging_cycle_bytes() -> bytes:
    """
    B1 = C1 * 3   (spectral radius = 9, diverges rapidly)
    C1 = B1 * 3
    Should not raise; should return (without hanging) even if not converged.
    """
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Sheet1"
    ws["B1"] = "=C1*3"
    ws["C1"] = "=B1*3"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


@pytest.fixture
def circular_with_downstream_bytes() -> bytes:
    """
    A1 = 100 (hardcoded)
    B1 = (A1 + C1) / 2      circular with C1
    C1 = B1 * 0.1            circular with B1
    D1 = B1 + 500            downstream of circular, non-circular
    Fixed point: B1 ≈ 52.6316, so D1 ≈ 552.6316
    """
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = 100
    ws["B1"] = "=(A1+C1)/2"
    ws["C1"] = "=B1*0.1"
    ws["D1"] = "=B1+500"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def test_self_reference_converges(self_ref_bytes):
    """Self-referential cell seeds to 0 and stays at 0."""
    workbook = parse_xlsx(self_ref_bytes)
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    a1 = next(c for c in workbook.all_cells if c.formula)
    assert result[a1.stable_id] == 0


def test_two_cell_cycle_converges(two_cell_cycle_bytes):
    """B1 and C1 in a cycle converge to the analytic fixed point."""
    workbook = parse_xlsx(two_cell_cycle_bytes)
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    b1 = next(c for c in workbook.all_cells if c.formula and "A1" in (c.formula or "") and "C1" in (c.formula or ""))
    c1 = next(c for c in workbook.all_cells if c.formula and "B1" in (c.formula or "") and "0.1" in (c.formula or ""))
    assert abs(result[b1.stable_id] - (1000 / 1.9)) < 0.1
    assert abs(result[c1.stable_id] - (100 / 1.9)) < 0.1


def test_interest_chain_converges(interest_chain_bytes):
    """3-cell interest chain converges to analytic fixed point."""
    workbook = parse_xlsx(interest_chain_bytes)
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    # A3 (EBT) should be ≈ 9090.91
    a3 = next(c for c in workbook.all_cells if c.formula and "A4" in (c.formula or "") and "A1" in (c.formula or ""))
    assert abs(result[a3.stable_id] - (10_000 / 1.1)) < 1.0


def test_diverging_cycle_does_not_crash(diverging_cycle_bytes):
    """Diverging circular chain completes without exception."""
    workbook = parse_xlsx(diverging_cycle_bytes)
    run_all(workbook)
    # Must not raise; return value may be large numbers or None
    result = evaluate_patch({}, workbook)
    assert isinstance(result, dict)


def test_downstream_of_circular_uses_converged_value(circular_with_downstream_bytes):
    """D1 = B1 + 500 where B1 is circular; D1 should use converged B1."""
    workbook = parse_xlsx(circular_with_downstream_bytes)
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    d1 = next(c for c in workbook.all_cells if c.formula and "500" in (c.formula or ""))
    b1 = next(c for c in workbook.all_cells if c.formula and "A1" in (c.formula or "") and "C1" in (c.formula or ""))
    b1_val = result[b1.stable_id]
    d1_val = result[d1.stable_id]
    assert b1_val is not None
    assert d1_val is not None
    assert abs(d1_val - (b1_val + 500)) < 0.1


# --- OFFSET-as-range-endpoint tests -----------------------------------------


def _single_sheet_bytes(cells: dict, title: str = "Sheet1") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for addr, val in cells.items():
        ws[addr] = val
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_offset_simple_negative_col():
    """E1 = SUM(A1:OFFSET(A1,0,2)) → A1:C1 = 1+2+3 = 6."""
    data = {
        "A1": 1, "B1": 2, "C1": 3, "D1": 4,
        "E1": "=SUM(A1:OFFSET(A1,0,2))",
    }
    workbook = parse_xlsx(_single_sheet_bytes(data))
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    e1 = next(c for c in workbook.all_cells if c.formula and "OFFSET" in c.formula)
    assert result[e1.stable_id] == 6


def test_offset_inside_average_with_min():
    """=AVERAGE(L322:OFFSET(L322,0,MIN(2,L5-1))) → L322:N322 → avg of 1,2,3 = 2.0"""
    data = {
        "L322": 1, "M322": 2, "N322": 3, "O322": 4, "P322": 5,
        "L5": 3,
        "A1": "=AVERAGE(L322:OFFSET(L322,0,MIN(2,L5-1)))",
    }
    workbook = parse_xlsx(_single_sheet_bytes(data))
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    a1 = next(c for c in workbook.all_cells if c.formula and "AVERAGE" in c.formula)
    assert abs(result[a1.stable_id] - 2.0) < 1e-9


def test_offset_with_bare_cell_ref():
    """=AVERAGE(L322:OFFSET(L322,0,L5)) where L5=2 → L322:N322 → avg of 10,20,30 = 20.0"""
    data = {
        "L322": 10, "M322": 20, "N322": 30,
        "L5": 2,
        "A1": "=AVERAGE(L322:OFFSET(L322,0,L5))",
    }
    workbook = parse_xlsx(_single_sheet_bytes(data))
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    a1 = next(c for c in workbook.all_cells if c.formula and "AVERAGE" in c.formula)
    assert abs(result[a1.stable_id] - 20.0) < 1e-9


def test_offset_out_of_range_returns_none():
    """=OFFSET(A1,0,-5) → new col would be -4 → caller returns None."""
    data = {
        "A1": 1,
        "B1": "=OFFSET(A1,0,-5)",
    }
    workbook = parse_xlsx(_single_sheet_bytes(data))
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    b1 = next(c for c in workbook.all_cells if c.formula)
    assert result[b1.stable_id] is None


def test_offset_with_sheet_name():
    """Cross-sheet OFFSET reference resolves correctly."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Data"
    ws1["A1"] = 10
    ws1["B1"] = 20
    ws1["C1"] = 30
    ws2 = wb.create_sheet("Calc")
    ws2["A1"] = "=SUM(Data!A1:OFFSET(Data!A1,0,2))"
    buf = io.BytesIO()
    wb.save(buf)
    workbook = parse_xlsx(buf.getvalue())
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    calc = next(c for c in workbook.all_cells if c.formula and "SUM" in c.formula)
    assert result[calc.stable_id] == 60


# --- Structured-reference (Table[[#This Row],[Col]]) tests ------------------


def test_structured_ref_this_row_resolves():
    """=MyTbl[[#This Row],[B_col]]+1 at C2 -> B2+1 = 11."""
    from openpyxl.worksheet.table import Table, TableColumn
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'], ws['B1'] = 'A_col', 'B_col'
    ws['A2'], ws['B2'] = 1, 10
    ws['A3'], ws['B3'] = 2, 20
    ws['C2'] = "=MyTbl[[#This Row],[B_col]]+1"
    tbl = Table(displayName='MyTbl', name='MyTbl', ref='A1:B3', tableColumns=[
        TableColumn(id=1, name='A_col'),
        TableColumn(id=2, name='B_col'),
    ])
    ws.add_table(tbl)
    buf = io.BytesIO()
    wb.save(buf)
    workbook = parse_xlsx(buf.getvalue())
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    c2 = next(c for c in workbook.all_cells if c.formula and "MyTbl" in c.formula)
    assert result[c2.stable_id] == 11


def test_structured_ref_unknown_table_returns_none():
    """Reference to undefined table -> result is None, no crash."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws['A1'] = 5
    ws['B1'] = "=NoSuchTable[[#This Row],[NoCol]]+1"
    buf = io.BytesIO()
    wb.save(buf)
    workbook = parse_xlsx(buf.getvalue())
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    b1 = next(c for c in workbook.all_cells if c.formula)
    assert result[b1.stable_id] is None


def test_structured_ref_cross_sheet():
    """Table on Sheet1, formula on Sheet2 references the table."""
    from openpyxl.worksheet.table import Table, TableColumn
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Data Sheet"
    ws1['A1'], ws1['B1'] = 'A_col', 'B_col'
    ws1['A2'], ws1['B2'] = 1, 100
    ws1['A3'], ws1['B3'] = 2, 200
    tbl = Table(displayName='CrossTbl', name='CrossTbl', ref='A1:B3', tableColumns=[
        TableColumn(id=1, name='A_col'),
        TableColumn(id=2, name='B_col'),
    ])
    ws1.add_table(tbl)
    ws2 = wb.create_sheet("Calc")
    # Formula at C2 referencing #This Row -> row 2 of Data Sheet -> B2 = 100
    ws2['C2'] = "=CrossTbl[[#This Row],[B_col]]*2"
    buf = io.BytesIO()
    wb.save(buf)
    workbook = parse_xlsx(buf.getvalue())
    run_all(workbook)
    result = evaluate_patch({}, workbook)
    c2 = next(c for c in workbook.all_cells if c.formula and "CrossTbl" in c.formula)
    assert result[c2.stable_id] == 200
