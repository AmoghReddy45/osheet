# osheet

An AI-native spreadsheet compiler. Upload any `.xlsx` — get back a structured, agent-readable workbook.

## Install

```bash
pip install osheet
```

## Quickstart

```python
import osheet

wb = osheet.load(open("model.xlsx", "rb").read())
print(wb.assumptions)           # detected input cells
print(wb.outputs)               # detected output metrics
result = wb.trace("metric.gross_margin")  # upstream dependencies
```

## Development

```bash
cd packages/osheet
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
pytest
```

## Project Structure

```
packages/
  osheet/          # Python library (the engine)
  osheet-app/      # Web app (FastAPI + Next.js)
benchmarks/        # Baseline vs osheet accuracy comparison
docs/              # Specs and implementation plans
```
