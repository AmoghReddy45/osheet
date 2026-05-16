"""Run both benchmarks and print a comparison table."""
from run_baseline import run_baseline
from run_osheet import run_osheet


def main() -> None:
    print("Running baseline benchmark...")
    baseline = run_baseline()
    print("\nRunning osheet benchmark...")
    osheet_result = run_osheet()

    print("\n" + "=" * 60)
    print(f"{'Metric':<30} {'Baseline':>12} {'osheet':>12}")
    print("-" * 60)
    print(f"{'Accuracy':<30} {baseline.accuracy:>11.1%} {osheet_result.accuracy:>11.1%}")
    print(f"{'Avg Latency (ms)':<30} {baseline.avg_latency_ms:>11.0f} {osheet_result.avg_latency_ms:>11.0f}")
    print(f"{'Questions Correct':<30} {sum(r.correct for r in baseline.results):>12} {sum(r.correct for r in osheet_result.results):>12}")
    print("=" * 60)

    print("\nPer-question breakdown:")
    for b, o in zip(baseline.results, osheet_result.results):
        b_mark = "✓" if b.correct else "✗"
        o_mark = "✓" if o.correct else "✗"
        print(f"  {b.question_id}: baseline={b_mark} ({b.latency_ms:.0f}ms)  osheet={o_mark} ({o.latency_ms:.0f}ms)")


if __name__ == "__main__":
    main()
