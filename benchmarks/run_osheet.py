"""osheet benchmark: load with library, give Claude the structured manifest + cells."""
import sys
import time
import anthropic
import osheet
from metrics import QUESTIONS, BenchmarkResult, QuestionResult, score_answer


def workbook_to_context(wb) -> str:
    lines = []
    lines.append(f"Workbook: {wb.manifest.sheet_count} sheets, {wb.manifest.table_count} tables")
    lines.append(f"\nAssumptions ({len(wb.assumptions)}):")
    for c in wb.assumptions:
        lines.append(f"  [{c.stable_id}] = {c.value}")
    lines.append(f"\nOutputs ({len(wb.outputs)}):")
    for c in wb.outputs:
        val = c.value if c.value is not None else f"formula: {c.formula}"
        lines.append(f"  [{c.stable_id}] = {val}")
    lines.append("\nFormula dependencies (sample):")
    for c in wb.all_cells:
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
