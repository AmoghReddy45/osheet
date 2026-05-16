"""osheet benchmark: load with library, give Claude the structured manifest + cells."""
import sys
import time
import anthropic
import osheet
from metrics import QUESTIONS, BenchmarkResult, QuestionResult, score_answer


def _label_for(c, cell_map: dict) -> str:
    """Find the nearest text label in the same row (look left up to 4 cols)."""
    for offset in range(1, 5):
        neighbor = cell_map.get((c.sheet_name, c.row, c.col - offset))
        if neighbor and isinstance(neighbor.value, str) and neighbor.value.strip():
            return neighbor.value.strip()
    return ""


def workbook_to_context(wb) -> str:
    all_cells = [c for s in wb.sheets for c in s.cells]
    cell_map = {(c.sheet_name, c.row, c.col): c for c in all_cells}

    lines = []
    lines.append(f"Workbook: {wb.manifest.sheet_count} sheets, {wb.manifest.table_count} tables")

    lines.append(f"\nAssumptions ({len(wb.assumptions)}):")
    for c in wb.assumptions:
        label = _label_for(c, cell_map)
        label_str = f" ({label})" if label else ""
        lines.append(f"  [{c.stable_id}]{label_str} = {c.value}")

    lines.append(f"\nOutputs ({len(wb.outputs)}):")
    for c in wb.outputs:
        label = _label_for(c, cell_map)
        label_str = f" ({label})" if label else ""
        val = c.value if c.value is not None else f"formula: {c.formula}"
        lines.append(f"  [{c.stable_id}]{label_str} = {val}")

    lines.append("\nFormula dependencies (sample):")
    for c in all_cells:
        if c.depends_on:
            lines.append(f"  {c.stable_id} depends on: {', '.join(c.depends_on[:3])}")
    return "\n".join(lines)


def run_osheet(xlsx_path: str = "benchmarks/dummy_financial_model.xlsx") -> BenchmarkResult:
    with open(xlsx_path, "rb") as f:
        data = f.read()

    wb = osheet.load(data)
    context = workbook_to_context(wb)
    client = anthropic.Anthropic()
    results = []

    for q in QUESTIONS:
        prompt = f"""Here is a structured AI-native workbook representation:

{context}

Question: {q['question']}

Answer concisely using the structured data above."""

        t0 = time.time()
        response = client.messages.create(
            model="claude-opus-4-7",
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        latency_ms = (time.time() - t0) * 1000
        answer = response.content[0].text

        result = QuestionResult(
            question_id=q["id"],
            question=q["question"],
            answer=answer,
            correct=score_answer(answer, q["expected_contains"]),
            latency_ms=latency_ms,
        )
        results.append(result)
        print(f"[osheet]   {q['id']}: {'✓' if result.correct else '✗'} ({latency_ms:.0f}ms)")

    return BenchmarkResult(label="osheet_structured", results=results)


if __name__ == "__main__":
    result = run_osheet()
    print(f"\nOsheet accuracy: {result.accuracy:.1%}  avg latency: {result.avg_latency_ms:.0f}ms")
