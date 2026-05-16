"""Baseline: dump xlsx as CSV text and ask Claude to answer questions about it."""
import sys
import time
import openpyxl
import anthropic
from metrics import QUESTIONS, BenchmarkResult, QuestionResult, score_answer


def xlsx_to_text(path: str) -> str:
    wb = openpyxl.load_workbook(path, data_only=True)
    parts = []
    for ws in wb.worksheets:
        parts.append(f"=== Sheet: {ws.title} ===")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None for v in row):
                parts.append("\t".join(str(v) if v is not None else "" for v in row))
    return "\n".join(parts)


def run_baseline(xlsx_path: str = "benchmarks/dummy_financial_model.xlsx") -> BenchmarkResult:
    spreadsheet_text = xlsx_to_text(xlsx_path)
    client = anthropic.Anthropic()
    results = []

    for q in QUESTIONS:
        prompt = f"""Here is a spreadsheet exported as text:

{spreadsheet_text}

Question: {q['question']}

Answer concisely."""

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
        print(f"[baseline] {q['id']}: {'✓' if result.correct else '✗'} ({latency_ms:.0f}ms)")

    return BenchmarkResult(label="baseline_raw_csv", results=results)


if __name__ == "__main__":
    result = run_baseline()
    print(f"\nBaseline accuracy: {result.accuracy:.1%}  avg latency: {result.avg_latency_ms:.0f}ms")
