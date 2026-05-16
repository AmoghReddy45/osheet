import io, openpyxl

def make_circ_model() -> bytes:
    """
    3-statement model with circular reference:
    - EBITDA → interest expense → net debt → interest expense (cycle)
    - Controlled by a CIRC_REF_SWITCH (standard LBO pattern)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"

    # Assumptions (no fills)
    ws["A1"] = "CIRC_REF_SWITCH"
    ws["B1"] = 1          # 1 = ON (circular), 0 = OFF
    ws["A2"] = "Revenue"
    ws["B2"] = 10_000_000
    ws["A3"] = "EBITDA_margin"
    ws["B3"] = 0.30
    ws["A4"] = "Debt"
    ws["B4"] = 5_000_000
    ws["A5"] = "Interest_rate"
    ws["B5"] = 0.08
    ws["A6"] = "Tax_rate"
    ws["B6"] = 0.25

    # Calculations
    ws["A8"] = "EBITDA"
    ws["B8"] = "=B2*B3"   # 3,000,000

    ws["A9"] = "Interest_expense"
    # The circular reference: if switch ON, use avg of opening/closing debt
    # Simplified: interest = debt * rate (closing debt depends on net income which depends on interest)
    ws["B9"] = "=IF(B1=1, B4*B5, B4*B5)"  # simplified circular (references B4 which is static for now)

    ws["A10"] = "EBT"
    ws["B10"] = "=B8-B9"

    ws["A11"] = "Tax"
    ws["B11"] = "=B10*B6"

    ws["A12"] = "Net_income"
    ws["B12"] = "=B10-B11"

    ws["A13"] = "Cash_sweep"
    ws["B13"] = "=B12*0.5"   # 50% of net income goes to debt repayment

    ws["A14"] = "Closing_debt"
    ws["B14"] = "=B4-B13"    # debt reduces by cash sweep

    # True circular: interest depends on closing debt, closing debt depends on net income, net income depends on interest
    # Override B9 with actual circular formula
    ws["B9"] = "=IF(B1=1, (B4+B14)/2*B5, B4*B5)"  # <-- THIS IS THE CIRCULAR REFERENCE
    # B14 = B4 - B13 = B4 - B12*0.5 = B4 - (B10-B11)*0.5 = ... which flows through B9

    ws["A15"] = "Check"
    ws["B15"] = "=B4-B14-B13"  # should = 0 if model checks out

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

if __name__ == "__main__":
    data = make_circ_model()
    path = "benchmarks/circ_ref_model.xlsx"
    with open(path, "wb") as f:
        f.write(data)
    print(f"Written {len(data):,} bytes → {path}")
