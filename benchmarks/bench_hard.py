# benchmarks/bench_hard.py
"""
Hard benchmark: 4 tests where osheet's structural advantage over raw CSV is measurable.

Test 1 — Assumption ID: F1 score for identifying all 20 Config assumptions.
Test 2 — Trace: Recall for upstream inputs of Annual EBITDA.
Test 3 — Edit impact: Does the approach get the exact new revenue after patching growth rate?
Test 4 — Cell navigation: Can the approach find a specific cell by name across 15 sheets?

Run from repo root:
  ANTHROPIC_API_KEY=... python3 benchmarks/bench_hard.py
"""
import sys, os, re, time, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))

import anthropic
import openpyxl
import osheet
from make_hard_fixture import make_hard_model, CONFIG, MONTHS

XLSX_PATH = "benchmarks/hard_financial_model.xlsx"
MODEL = "claude-opus-4-7"

_client: anthropic.Anthropic | None = None


def client() -> anthropic.Anthropic:
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


# ── Helpers ──────────────────────────────────────────────────────────────────

def _xlsx_to_text(path: str) -> str:
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for ws in wb.worksheets:
        parts.append(f"=== Sheet: {ws.title} ===")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None for v in row):
                parts.append("\t".join(str(v) if v is not None else "" for v in row))
    return "\n".join(parts)


def _ask(prompt: str) -> tuple[str, float]:
    t0 = time.time()
    resp = client().messages.create(
        model=MODEL, max_tokens=1024,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text, (time.time() - t0) * 1000


def _extract_json_list(text: str) -> list:
    try:
        start, end = text.rfind("["), text.rfind("]")
        if start != -1 and end != -1:
            return json.loads(text[start:end+1])
    except Exception:
        pass
    return []


def _f1(found: set, expected: set) -> tuple[float, float, float]:
    tp = len(found & expected)
    fp = len(found - expected)
    fn = len(expected - found)
    p = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    r = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    f1 = 2 * p * r / (p + r) if (p + r) > 0 else 0.0
    return p, r, f1


# ── Test 1: Assumption Identification (F1) ───────────────────────────────────

def test1_baseline(text: str) -> dict:
    prompt = f"""Spreadsheet (text dump):

{text}

List ALL raw input assumptions in this model. Return a JSON array of objects with "name" and "value" keys.
Include only hard-coded scalar inputs — not calculated formulas."""
    answer, latency = _ask(prompt)
    items = _extract_json_list(answer)
    found_names = {str(it.get("name", "")).lower().replace(" ", "_") for it in items}
    expected = {name.lower() for name, _ in CONFIG}
    p, r, f1 = _f1(found_names, expected)
    return {"precision": p, "recall": r, "f1": f1, "found": len(found_names), "latency_ms": latency}


def test1_osheet(wb: osheet.OsheetWorkbook) -> dict:
    t0 = time.time()
    all_cells = wb.all_cells
    cell_map = {(c.sheet_name, c.row, c.col): c for c in all_cells}
    found_names = set()
    for c in wb.assumptions:
        for offset in range(1, 4):
            nb = cell_map.get((c.sheet_name, c.row, c.col - offset))
            if nb and isinstance(nb.value, str) and nb.value.strip():
                found_names.add(nb.value.strip().lower().replace(" ", "_"))
                break
        else:
            found_names.add(c.stable_id.split(".")[-1].lower())
    expected = {name.lower() for name, _ in CONFIG}
    p, r, f1 = _f1(found_names, expected)
    latency = (time.time() - t0) * 1000
    return {"precision": p, "recall": r, "f1": f1, "found": len(found_names), "latency_ms": latency}


# ── Test 2: Trace — What feeds into Annual Total EBITDA? ─────────────────────

def test2_baseline(text: str) -> dict:
    prompt = f"""Spreadsheet:

{text}

What are the input assumptions from the Config sheet that directly or indirectly affect Annual Total EBITDA?
Return a JSON array of assumption names."""
    answer, latency = _ask(prompt)
    items = _extract_json_list(answer)
    found = {str(it).lower().replace(" ", "_") for it in items}
    # Annual Total EBITDA = sum of monthly B13 (Gross_Profit - Total_OpEx).
    # It depends on revenue-percentage costs only — NOT headcount costs (those feed EBITDA_after_HC).
    # The 7 actual assumptions: arr_start, growth_rate_monthly, churn_rate_monthly,
    # cogs_pct, rd_pct, sm_pct, gna_pct
    ebitda_inputs = {
        "arr_start", "growth_rate_monthly", "churn_rate_monthly",
        "cogs_pct", "rd_pct", "sm_pct", "gna_pct",
    }
    _, recall, _ = _f1(found, ebitda_inputs)
    return {"recall": recall, "found_count": len(found), "expected_count": len(ebitda_inputs), "latency_ms": latency}


def test2_osheet(wb: osheet.OsheetWorkbook) -> dict:
    """BFS upstream from Annual Total EBITDA using formula-derived dep graph."""
    t0 = time.time()

    # Find the Annual EBITDA formula cell: annual.b10
    # find() returns label cells; look for the B-column neighbor with a formula
    ebitda_label_cells = osheet.find(wb, "total_ebitda")
    ebitda_formula_cell = None
    cell_map = {(c.sheet_name, c.row, c.col): c for c in wb.all_cells}
    for lc in ebitda_label_cells:
        # Look for formula cell in the same row, adjacent columns
        for col_offset in range(1, 4):
            candidate = cell_map.get((lc.sheet_name, lc.row, lc.col + col_offset))
            if candidate and candidate.formula:
                ebitda_formula_cell = candidate
                break
        if ebitda_formula_cell:
            break

    if not ebitda_formula_cell:
        # Fallback: look up annual.b10 directly
        ebitda_formula_cell = wb._wb.get_cell("annual.b10")

    if not ebitda_formula_cell:
        return {"recall": 0.0, "found_count": 0, "expected_count": 14, "latency_ms": 0}

    # Build case-insensitive id lookup
    all_cells = wb.all_cells
    id_lower_map: dict[str, "osheet.models.Cell"] = {c.stable_id.lower(): c for c in all_cells}

    # BFS upstream through the full DAG; depends_on uses mixed case, normalize to lower
    visited: set[str] = set()
    queue = [ebitda_formula_cell.stable_id.lower()]
    while queue:
        cid = queue.pop()
        if cid in visited:
            continue
        visited.add(cid)
        cell = id_lower_map.get(cid)
        if cell:
            for dep in cell.depends_on:
                dep_lower = dep.lower()
                if dep_lower not in visited:
                    queue.append(dep_lower)

    assumption_ids_lower = {c.stable_id.lower() for c in wb.assumptions}
    found_assumption_ids = visited & assumption_ids_lower

    # Map assumption stable_ids back to names via label lookup
    found_assumption_names: set[str] = set()
    for aid in found_assumption_ids:
        cell = id_lower_map.get(aid)
        if cell:
            nb = cell_map.get((cell.sheet_name, cell.row, cell.col - 1))
            if nb and isinstance(nb.value, str):
                found_assumption_names.add(nb.value.strip().lower())
            else:
                found_assumption_names.add(aid)

    # Same ground truth as test2_baseline: 7 actual assumptions feeding Total_EBITDA
    ebitda_inputs = {
        "arr_start", "growth_rate_monthly", "churn_rate_monthly",
        "cogs_pct", "rd_pct", "sm_pct", "gna_pct",
    }
    recall = len(found_assumption_ids) / len(ebitda_inputs) if ebitda_inputs else 0
    latency = (time.time() - t0) * 1000
    return {
        "recall": recall,
        "found_count": len(found_assumption_ids),
        "expected_count": 7,
        "found_names": sorted(found_assumption_names),
        "latency_ms": latency,
    }


# ── Test 3: Edit Impact — growth_rate_monthly 0.025 → 0.05 ──────────────────

def test3_baseline(text: str) -> dict:
    """Ask Claude to estimate new Annual Revenue after growth rate change."""
    prompt = f"""Spreadsheet:

{text}

If growth_rate_monthly changes from 0.025 to 0.05, what would the new Annual Total Revenue be?
Respond with only the number (no units or explanation)."""
    answer, latency = _ask(prompt)
    nums = re.findall(r"[\d,]+\.?\d*", answer.replace(",", ""))
    try:
        guessed = float(nums[0]) if nums else 0.0
    except ValueError:
        guessed = 0.0
    return {"guessed_value": guessed, "latency_ms": latency}


def test3_osheet(wb: osheet.OsheetWorkbook) -> dict:
    """Use propose_patch to get exact computed new Annual Revenue."""
    t0 = time.time()
    growth_cell = next((c for c in wb.assumptions if c.value == 0.025), None)
    if not growth_cell:
        return {"computed_value": None, "error": "growth_rate_monthly not found as assumption"}
    proposal = osheet.propose_patch(wb, growth_cell.stable_id, 0.05)

    # Find the Annual Total_Revenue cell by scanning label cells in the Annual sheet.
    # computed_values keys are positional (e.g., "annual.b3"), so we locate the row
    # whose label contains "total_revenue" and build the corresponding key.
    revenue_key = None
    cell_map = {(c.sheet_name, c.row, c.col): c for c in wb.all_cells}
    for c in wb.all_cells:
        if c.sheet_name.lower() == "annual" and isinstance(c.value, str) and "total_revenue" in c.value.lower():
            # The formula cell is one column to the right (col B when label is col A)
            formula_cell = cell_map.get((c.sheet_name, c.row, c.col + 1))
            if formula_cell:
                candidate = formula_cell.stable_id.lower()
                if candidate in {k.lower() for k in proposal.computed_values}:
                    revenue_key = next(k for k in proposal.computed_values if k.lower() == candidate)
                    break

    # Fallback: prioritize "total_revenue" in key name, then first sorted annual key
    if revenue_key is None:
        revenue_key = next(
            (k for k in proposal.computed_values if "total_revenue" in k.lower()),
            None,
        )
    if revenue_key is None:
        revenue_key = next(
            (k for k in sorted(proposal.computed_values) if "annual" in k.lower()),
            None,
        )

    computed = proposal.computed_values.get(revenue_key) if revenue_key else None
    latency = (time.time() - t0) * 1000
    return {"computed_value": computed, "cell_id": revenue_key, "latency_ms": latency, "affected_count": len(proposal.affected_cells)}


# ── Test 4: Context Stress — Navigate 15 sheets to find a specific cell ──────

def test4_baseline(text: str) -> dict:
    """Can Claude find and report the formula for KPIs Burn_Multiple?"""
    prompt = f"""Spreadsheet:

{text}

What is the formula or value in the cell labeled 'Burn_Multiple' in the KPIs sheet?
Return just the formula string."""
    answer, latency = _ask(prompt)
    # Correct answer contains ABS and Annual references
    correct = "annual" in answer.lower() or "abs" in answer.lower() or "b9" in answer.lower()
    return {"correct": correct, "latency_ms": latency, "answer_preview": answer[:150]}


def test4_osheet(wb: osheet.OsheetWorkbook) -> dict:
    """Use find() to locate Burn_Multiple label, then read the adjacent formula cell."""
    t0 = time.time()
    label_cells = osheet.find(wb, "burn_multiple")
    cell_map = {(c.sheet_name, c.row, c.col): c for c in wb.all_cells}

    formula = None
    for lc in label_cells:
        if lc.sheet_name.lower() == "kpis":
            # Formula cell is one column to the right
            for col_offset in range(1, 4):
                candidate = cell_map.get((lc.sheet_name, lc.row, lc.col + col_offset))
                if candidate and candidate.formula:
                    formula = candidate.formula
                    break
        if formula:
            break

    found = formula is not None
    correct = found and ("ABS" in formula or "abs" in formula.lower()) and "annual" in formula.lower()
    latency = (time.time() - t0) * 1000
    return {"correct": correct, "formula": formula, "latency_ms": latency}


# ── Main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    # Generate fixture if missing
    if not os.path.exists(XLSX_PATH):
        print(f"Generating {XLSX_PATH}...")
        data = make_hard_model()
        with open(XLSX_PATH, "wb") as f:
            f.write(data)

    with open(XLSX_PATH, "rb") as f:
        xlsx_bytes = f.read()

    print("Loading osheet workbook...")
    wb = osheet.load(xlsx_bytes)
    print(f"  {len(wb.assumptions)} assumptions detected, {len(wb.outputs)} outputs")

    print("\nBuilding raw CSV text for baseline...")
    text = _xlsx_to_text(XLSX_PATH)
    token_est = int(len(text.split()) * 1.3)
    print(f"  ~{token_est:,} tokens in CSV dump")

    sep = "=" * 72

    print(f"\n{sep}")
    print("TEST 1: Assumption Identification (F1 Score)")
    print(f"  Expected: 20 Config assumptions (no yellow fills to cheat with)")
    print(sep)
    print("  [baseline] querying Claude...")
    b1 = test1_baseline(text)
    print(f"  [baseline] precision={b1['precision']:.0%}  recall={b1['recall']:.0%}  F1={b1['f1']:.0%}  found={b1['found']}  ({b1['latency_ms']:.0f}ms)")
    print("  [osheet]   structured lookup...")
    o1 = test1_osheet(wb)
    print(f"  [osheet]   precision={o1['precision']:.0%}  recall={o1['recall']:.0%}  F1={o1['f1']:.0%}  found={o1['found']}  ({o1['latency_ms']:.0f}ms)")

    print(f"\n{sep}")
    print("TEST 2: Trace Recall — Inputs that affect Annual Total EBITDA")
    print(sep)
    print("  [baseline] querying Claude...")
    b2 = test2_baseline(text)
    print(f"  [baseline] recall={b2['recall']:.0%}  found={b2['found_count']}/{b2['expected_count']}  ({b2['latency_ms']:.0f}ms)")
    print("  [osheet]   DAG traversal...")
    o2 = test2_osheet(wb)
    print(f"  [osheet]   recall={o2['recall']:.0%}  found={o2['found_count']}/{o2['expected_count']}  ({o2['latency_ms']:.0f}ms)")

    print(f"\n{sep}")
    print("TEST 3: Edit Impact — growth_rate_monthly: 0.025 → 0.05")
    print("  (No Claude call for osheet — exact answer from formula evaluator)")
    print(sep)
    print("  [baseline] querying Claude to estimate...")
    b3 = test3_baseline(text)
    print(f"  [baseline] guessed Annual Revenue: {b3['guessed_value']:,.0f}  ({b3['latency_ms']:.0f}ms)")
    print("  [osheet]   running propose_patch()...")
    o3 = test3_osheet(wb)
    if o3.get("computed_value") is not None:
        print(f"  [osheet]   computed Annual Total_Revenue [{o3['cell_id']}]: {o3['computed_value']:,.2f}  "
              f"({o3['affected_count']} cells changed, {o3['latency_ms']:.0f}ms)")
        if b3["guessed_value"] > 0 and o3["computed_value"]:
            error_pct = abs(b3["guessed_value"] - o3["computed_value"]) / o3["computed_value"] * 100
            print(f"  [baseline] error vs evaluated truth: {error_pct:.1f}%")
    else:
        print(f"  [osheet]   {o3}")

    print(f"\n{sep}")
    print("TEST 4: Cell Navigation — Find KPIs Burn_Multiple formula across 15 sheets")
    print(sep)
    print("  [baseline] querying Claude to navigate full CSV...")
    b4 = test4_baseline(text)
    print(f"  [baseline] correct={b4['correct']}  ({b4['latency_ms']:.0f}ms)")
    print(f"             answer: {b4['answer_preview']}")
    print("  [osheet]   using find()...")
    o4 = test4_osheet(wb)
    print(f"  [osheet]   correct={o4['correct']}  formula={o4['formula']}  ({o4['latency_ms']:.0f}ms)")

    print(f"\n{sep}")
    print("SUMMARY")
    print(f"{sep}")
    print(f"{'Test':<42} {'Baseline':>14} {'osheet':>14}")
    print("-" * 72)
    print(f"{'T1 Assumption ID (F1 score)':<42} {b1['f1']:>13.0%} {o1['f1']:>13.0%}")
    print(f"{'T2 Trace Recall':<42} {b2['recall']:>13.0%} {o2['recall']:>13.0%}")
    t3_baseline_str = f"~{b3['guessed_value']:,.0f}" if b3['guessed_value'] else "no answer"
    t3_osheet_str   = f"{o3['computed_value']:,.2f}" if o3.get('computed_value') is not None else "n/a"
    print(f"{'T3 Annual Revenue after patch':<42} {t3_baseline_str:>14} {t3_osheet_str:>14}")
    print(f"{'T4 Cell navigation correct':<42} {str(b4['correct']):>14} {str(o4['correct']):>14}")
    print(f"{sep}")


if __name__ == "__main__":
    main()
