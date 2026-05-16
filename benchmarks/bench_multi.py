"""Run N iterations of each benchmark and report averaged results."""
import sys
import statistics
from run_baseline import run_baseline
from run_osheet import run_osheet
from metrics import BenchmarkResult


def avg_result(runs: list[BenchmarkResult], label: str) -> dict:
    accuracies = [r.accuracy for r in runs]
    latencies = [r.avg_latency_ms for r in runs]
    per_q: dict[str, list[bool]] = {}
    for r in runs:
        for qr in r.results:
            per_q.setdefault(qr.question_id, []).append(qr.correct)
    return {
        "label": label,
        "accuracy_mean": statistics.mean(accuracies),
        "accuracy_stdev": statistics.stdev(accuracies) if len(accuracies) > 1 else 0,
        "latency_mean_ms": statistics.mean(latencies),
        "per_q": {qid: sum(vals) / len(vals) for qid, vals in per_q.items()},
    }


def main(n: int = 3):
    print(f"Running {n} iterations of each benchmark...\n")

    baseline_runs = []
    for i in range(n):
        print(f"=== Baseline run {i+1}/{n} ===")
        baseline_runs.append(run_baseline())

    osheet_runs = []
    for i in range(n):
        print(f"\n=== osheet run {i+1}/{n} ===")
        osheet_runs.append(run_osheet())

    b = avg_result(baseline_runs, "baseline")
    o = avg_result(osheet_runs, "osheet")

    print("\n" + "=" * 70)
    print(f"{'Metric':<35} {'Baseline':>15} {'osheet':>15}")
    print("-" * 70)
    print(f"{'Accuracy (mean)':<35} {b['accuracy_mean']:>14.1%} {o['accuracy_mean']:>14.1%}")
    print(f"{'Accuracy (stdev)':<35} {b['accuracy_stdev']:>14.1%} {o['accuracy_stdev']:>14.1%}")
    print(f"{'Avg Latency ms (mean)':<35} {b['latency_mean_ms']:>14.0f} {o['latency_mean_ms']:>14.0f}")
    print("=" * 70)
    print("\nPer-question accuracy (across all runs):")
    for qid in sorted(b["per_q"]):
        bq = b["per_q"][qid]
        oq = o["per_q"][qid]
        print(f"  {qid}: baseline={bq:.0%}  osheet={oq:.0%}")


if __name__ == "__main__":
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 3
    main(n)
