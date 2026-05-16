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
