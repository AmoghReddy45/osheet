"""Scoring utilities for benchmark runs."""
from dataclasses import dataclass


QUESTIONS = [
    {
        "id": "q1",
        "question": "What is the churn rate assumption used in this model?",
        "expected_contains": ["0.04", "4%", "4 percent"],
    },
    {
        "id": "q2",
        "question": "What is the gross margin percentage?",
        "expected_contains": ["0.72", "72%", "72 percent"],
    },
    {
        "id": "q3",
        "question": "Which cells or values are inputs/assumptions that I can change?",
        "expected_contains": ["churn", "growth", "margin", "assumption"],
    },
    {
        "id": "q4",
        "question": "What is the total ARR at the end of the year?",
        "expected_contains": ["arr", "revenue", "cumulative"],
    },
    {
        "id": "q5",
        "question": "If I change the growth rate to 20%, which outputs would change?",
        "expected_contains": ["arr", "revenue", "profit", "gross"],
    },
]


@dataclass
class QuestionResult:
    question_id: str
    question: str
    answer: str
    correct: bool
    latency_ms: float


@dataclass
class BenchmarkResult:
    label: str
    results: list[QuestionResult]

    @property
    def accuracy(self) -> float:
        if not self.results:
            return 0.0
        return sum(r.correct for r in self.results) / len(self.results)

    @property
    def avg_latency_ms(self) -> float:
        if not self.results:
            return 0.0
        return sum(r.latency_ms for r in self.results) / len(self.results)


def score_answer(answer: str, expected_contains: list[str]) -> bool:
    answer_lower = answer.lower()
    return any(e.lower() in answer_lower for e in expected_contains)
