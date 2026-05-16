# osheet — Design Spec

**Date:** 2026-05-15  
**Status:** Approved  
**Repo:** https://github.com/AmoghReddy45/osheet

---

## What We're Building

An AI-native spreadsheet compiler. Upload any `.xlsx` file — get back a structured, agent-readable workbook with a clean Python API and an annotated `.xlsx` output.

> "Make spreadsheets readable by AI agents without making them weird for humans."

---

## Core Idea

```
Input:  messy_model.xlsx
Output: workbook (Python API)  +  model.ai.xlsx  +  model.osheet
```

The human sees the same spreadsheet. The AI sees a fully structured object model with stable IDs, typed tables, a formula dependency graph, and safe editing primitives.

---

## Users

**Primary:** Developers building AI agents that need to work with spreadsheets (LangChain, AutoGPT, custom agents).  
**Secondary:** Analysts and financial modelers who want AI to work reliably on their Excel files.

If we build the Python API well, both use cases are served.

---

## Architecture

### Monorepo layout

```
osheet/
  packages/
    osheet/          # Python library (the engine)
    osheet-app/      # Web app (FastAPI backend + Next.js frontend)
  benchmarks/        # Baseline vs. osheet comparison suite
  fixtures/          # Sample .xlsx files for testing
  docs/
```

### Library pipeline

Three sequential stages:

```
.xlsx
  └─▶ Parser        (openpyxl → raw cell/formula/style model)
        └─▶ Analyzer  (5 semantic inference passes)
              └─▶ Emitter (Workbook object + .xlsx + .osheet)
```

### Web app

- **Backend:** FastAPI — thin wrapper around the library. Three endpoints: `POST /convert`, `GET /result/{job_id}`, `GET /download/{job_id}/{file}`.
- **Frontend:** Next.js — three views: Upload, Inspector, Diff (stretch).

---

## Data Model

```
Workbook
├── sheets: List[Sheet]
│   ├── id: str                  # stable: "sheet.revenue_model"
│   ├── name: str
│   └── tables: List[Table]
│       ├── id: str              # "table.monthly_revenue"
│       ├── range: str           # "B4:N16"
│       ├── columns: List[Column]
│       │   ├── name: str
│       │   └── dtype: Literal["number","text","date","formula","unknown"]
│       └── cells: List[Cell]
│           ├── stable_id: str   # "metric.gross_margin.fy2026"
│           ├── role: Literal["assumption","output","intermediate","label","unknown"]
│           ├── formula: str | None
│           ├── value: Any
│           └── depends_on: List[str]
├── assumptions: List[Cell]      # flat index
├── outputs: List[Cell]          # flat index
├── formula_graph: DAG
└── manifest: Manifest           # confidence scores, warnings, conversion metadata
```

### .osheet format

A zip package:
```
workbook.json        # metadata + manifest
sheets.json          # Sheet + Table + Cell tree
formula_graph.json   # adjacency list DAG
styles.json          # visual formatting (for roundtrip)
original.xlsx        # embedded original
```

---

## Python API

```python
import osheet

wb = osheet.load("model.xlsx")

# Navigation
wb.assumptions                        # → List[Cell]
wb.outputs                            # → List[Cell]
wb.sheets                             # → List[Sheet]
wb.find("gross margin")               # → List[Cell]  (fuzzy label search)
wb.trace("metric.gross_margin.fy2026") # → TraceResult (upstream + downstream)

# Safe agent editing
proposal = wb.propose_patch("assumption.churn_rate.monthly", 0.06)
proposal.affected_cells               # → List[Cell] with new values
proposal.diff                         # → human-readable diff
wb.apply_patch(proposal)              # → mutates workbook

# Export
wb.export_xlsx() → bytes              # annotated .xlsx
wb.export_osheet() → bytes            # .osheet zip
wb.manifest.warnings                  # → List[Warning]
```

---

## Analyzer — Five Passes

**Pass 1 — Table detection**  
Finds contiguous rectangular regions with a header row. Uses formatting density, explicit Excel Table objects, and financial-model heuristics. Low-confidence detections flagged in manifest.

**Pass 2 — Column typing**  
Samples values per column → infers `dtype`. Extracts formula patterns for formula columns (e.g. `=B{r}*C{r}` repeated).

**Pass 3 — Cell role classification**  
Classifies each cell as `assumption` / `output` / `intermediate` / `label` / `unknown` using: formula presence, in/out-degree in dependency graph, fill color heuristics (yellow = assumption convention), positional heuristics.

**Pass 4 — Formula dependency graph**  
Parses every formula (regex + grammar) to build a DAG. Handles `SUM`, `IF`, `VLOOKUP`, named ranges, cross-sheet references. Resolves all references to stable IDs.

**Pass 5 — Stable ID assignment**  
Deterministic, human-readable IDs: `{sheet}.{table}.{column}.{row_label}`. Falls back to `{sheet}.{col}{row}`. Stable across re-conversions of the same file.

---

## Web App — Inspector UI

**Design language:** Geist + GeistMono fonts, `#171717` dark canvas, minimal chromatic accent.  
**Reference:** `/Users/amoghreddy/Downloads/Untitled Sheet__2026-05-15_20-52-58.html` (structural + component reference).

**Three-panel layout:**
- Left: sheet/table browser + legend
- Center: summary stat bar + cell table (stable ID, role badge, value, confidence)
- Right: cell detail (formula, dependency graph, agent note, trace/patch actions)

**Views:**
1. Upload — drag-and-drop, privacy note
2. Inspector — main output view (mockup finalized in brainstorm)
3. Diff — patch simulation (stretch goal)

---

## Error Handling

- Uncertain detections → `low_confidence` flag in manifest, not a failure.
- Malformed / password-protected xlsx → clean error with `reason` string.
- Agent-facing warnings exposed via `wb.manifest.warnings`.
- No stack traces surfaced to users.

---

## Testing

**Unit tests (pytest)**  
Each analyzer pass in isolation. One fixture per edge case: empty sheet, formula-only, cross-sheet refs, named ranges, pivot tables, password-protected (expected clean error).

**Integration tests**  
End-to-end: `.xlsx` in → `Workbook` out. Assertions on table count, assumption count, graph node count, stable ID format. Real-world-style financial model fixtures.

**API contract tests**  
Full Python API surface tested against known inputs/outputs. These protect agent consumers from regressions.

---

## Benchmarking Suite

**Baseline:** Run Claude directly on a raw dummy `.xlsx` (via CSV text dump). Ask navigation/analysis questions. Record accuracy + latency.

**Post-conversion:** Run same questions on the `osheet` Workbook API output. Record accuracy + latency.

**Metrics:** assumption detection accuracy, output cell identification accuracy, formula trace correctness, query latency, agent edit safety (no silent breakage).

---

## Tech Stack

| Layer | Choice | Reason |
|---|---|---|
| Library language | Python 3.11+ | Ecosystem fit; openpyxl, networkx, pydantic |
| xlsx parsing | openpyxl | Full Open XML support |
| Formula parsing | formulas / custom grammar | Handles cross-sheet refs |
| Graph | networkx | DAG traversal, shortest path |
| Data validation | pydantic v2 | Fast, typed models |
| Web backend | FastAPI | Shares Python with library |
| Web frontend | Next.js 14 (App Router) | Best-in-class DX |
| Styling | Tailwind CSS + Geist | Matches design language |
| Testing | pytest + httpx | Standard Python |
| CI | GitHub Actions | Already on GitHub |

---

## MVP Scope (v1)

In scope:
- Parser + all 5 analyzer passes
- Full Python API
- .osheet export format
- Annotated .xlsx export
- FastAPI backend
- Next.js Upload + Inspector views
- Unit + integration + contract tests
- Benchmarking suite

Out of scope for v1:
- MCP adapter (v2)
- Full workbook recompilation / layout restructuring (v2)
- Diff/patch UI view (stretch)
- Authentication / multi-user
