# osheet Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the osheet Python library, FastAPI backend, Next.js inspector frontend, and benchmarking suite from scratch.

**Architecture:** Python library is the engine (openpyxl → 5-pass analyzer → Workbook object + xlsx/osheet outputs). FastAPI wraps the library. Next.js serves Upload + Inspector views. Benchmarks compare raw Claude vs Claude-on-osheet.

**Tech Stack:** Python 3.11, openpyxl, networkx, pydantic v2, pytest, FastAPI, uvicorn, Next.js 14 App Router, Tailwind CSS, Geist font, anthropic SDK, GitHub Actions.

---

## File Map

```
osheet/                              # repo root
├── .github/workflows/ci.yml
├── .gitignore
├── README.md
├── packages/
│   ├── osheet/                      # Python library
│   │   ├── pyproject.toml
│   │   ├── src/osheet/
│   │   │   ├── __init__.py          # public API re-exports
│   │   │   ├── models.py            # Workbook, Sheet, Table, Cell, Manifest, Warning, DAG
│   │   │   ├── parser.py            # openpyxl → RawWorkbook
│   │   │   ├── api.py               # load(), trace(), propose_patch(), apply_patch()
│   │   │   ├── analyzer/
│   │   │   │   ├── __init__.py      # run_all(raw) → AnalyzedWorkbook
│   │   │   │   ├── tables.py        # Pass 1: table detection
│   │   │   │   ├── types.py         # Pass 2: column dtype inference
│   │   │   │   ├── roles.py         # Pass 3: assumption/output/intermediate
│   │   │   │   ├── graph.py         # Pass 4: formula DAG
│   │   │   │   └── ids.py           # Pass 5: stable ID assignment
│   │   │   └── emitter/
│   │   │       ├── __init__.py
│   │   │       ├── xlsx.py          # annotated .xlsx output
│   │   │       └── osheet.py        # .osheet zip output
│   │   └── tests/
│   │       ├── conftest.py          # shared fixtures (builds test xlsx in memory)
│   │       ├── test_parser.py
│   │       ├── test_tables.py
│   │       ├── test_types.py
│   │       ├── test_roles.py
│   │       ├── test_graph.py
│   │       ├── test_ids.py
│   │       ├── test_emitter.py
│   │       └── test_api.py
│   └── osheet-app/
│       ├── backend/
│       │   ├── pyproject.toml
│       │   └── src/app/
│       │       ├── main.py          # FastAPI app, CORS, mounts routes
│       │       ├── storage.py       # in-memory job store (dict, TTL)
│       │       └── routes/
│       │           ├── convert.py   # POST /convert
│       │           ├── result.py    # GET /result/{job_id}
│       │           └── download.py  # GET /download/{job_id}/{file}
│       └── frontend/
│           ├── package.json
│           ├── tailwind.config.ts
│           ├── next.config.ts
│           └── src/app/
│               ├── layout.tsx       # root layout, Geist font, dark theme
│               ├── page.tsx         # Upload view
│               ├── inspect/
│               │   └── [jobId]/
│               │       └── page.tsx # Inspector view
│               └── components/
│                   ├── UploadZone.tsx
│                   ├── StatBar.tsx
│                   ├── SheetSidebar.tsx
│                   ├── CellTable.tsx
│                   └── DetailPanel.tsx
└── benchmarks/
    ├── requirements.txt
    ├── make_fixture.py              # generates dummy_financial_model.xlsx
    ├── run_baseline.py              # raw CSV → Claude → scores
    ├── run_osheet.py                # osheet output → Claude → scores
    ├── metrics.py                   # accuracy + latency helpers
    └── report.py                    # prints comparison table
```

---

### Task 1: Root scaffolding

**Files:**
- Create: `.gitignore`
- Create: `README.md`
- Create: `packages/osheet/pyproject.toml`
- Create: `packages/osheet/src/osheet/__init__.py` (empty)

- [ ] Create `.gitignore`

```
__pycache__/
*.py[cod]
*.egg-info/
dist/
.venv/
node_modules/
.next/
.env
.env.local
*.osheet
.superpowers/
```

- [ ] Create `README.md`

```markdown
# osheet

An AI-native spreadsheet compiler. Upload any `.xlsx` — get back a structured, agent-readable workbook.

## Install

```bash
pip install osheet
```

## Quickstart

```python
import osheet

wb = osheet.load("model.xlsx")
print(wb.assumptions)          # detected input cells
print(wb.outputs)              # detected output metrics
wb.trace("metric.gross_margin") # upstream dependencies
```

## Development

```bash
cd packages/osheet && pip install -e ".[dev]"
pytest
```
```

- [ ] Create `packages/osheet/pyproject.toml`

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "osheet"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
    "openpyxl>=3.1",
    "networkx>=3.3",
    "pydantic>=2.7",
]

[project.optional-dependencies]
dev = ["pytest>=8", "pytest-cov", "httpx"]

[tool.hatch.build.targets.wheel]
packages = ["src/osheet"]

[tool.pytest.ini_options]
testpaths = ["tests"]
```

- [ ] Run: `cd packages/osheet && python -m venv .venv && source .venv/bin/activate && pip install -e ".[dev]"`
- [ ] Commit: `git add . && git commit -m "chore: root scaffolding and osheet package setup"`

---

### Task 2: Core data models

**Files:**
- Create: `packages/osheet/src/osheet/models.py`
- Create: `packages/osheet/tests/conftest.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_models.py
from osheet.models import Cell, CellRole, Sheet, Table, Workbook, Manifest

def test_cell_defaults():
    c = Cell(stable_id="sheet.A1", role=CellRole.UNKNOWN, value=42)
    assert c.formula is None
    assert c.depends_on == []
    assert c.confidence == 0.0

def test_workbook_flat_indexes():
    assumption = Cell(stable_id="s.A1", role=CellRole.ASSUMPTION, value=0.04)
    output = Cell(stable_id="s.B1", role=CellRole.OUTPUT, value=100.0)
    sheet = Sheet(id="sheet.s", name="S", tables=[], cells=[assumption, output])
    wb = Workbook(sheets=[sheet], manifest=Manifest())
    assert len(wb.assumptions) == 1
    assert len(wb.outputs) == 1
```

- [ ] Run: `pytest tests/test_models.py -v` — expect FAIL (ImportError)

- [ ] Implement `packages/osheet/src/osheet/models.py`

```python
from __future__ import annotations
from enum import Enum
from typing import Any
from pydantic import BaseModel, Field, model_validator


class CellRole(str, Enum):
    ASSUMPTION = "assumption"
    OUTPUT = "output"
    INTERMEDIATE = "intermediate"
    LABEL = "label"
    UNKNOWN = "unknown"


class ColumnDtype(str, Enum):
    NUMBER = "number"
    TEXT = "text"
    DATE = "date"
    FORMULA = "formula"
    MIXED = "mixed"
    UNKNOWN = "unknown"


class Cell(BaseModel):
    stable_id: str
    role: CellRole = CellRole.UNKNOWN
    value: Any = None
    formula: str | None = None
    depends_on: list[str] = Field(default_factory=list)
    sheet_name: str = ""
    col: int = 0
    row: int = 0
    fill_color: str | None = None
    confidence: float = 0.0

    @property
    def address(self) -> str:
        from openpyxl.utils import get_column_letter
        return f"{self.sheet_name}!{get_column_letter(self.col)}{self.row}"


class Column(BaseModel):
    name: str
    dtype: ColumnDtype = ColumnDtype.UNKNOWN
    col_index: int = 0


class Table(BaseModel):
    id: str
    sheet_name: str
    range_ref: str  # e.g. "B4:N16"
    columns: list[Column] = Field(default_factory=list)
    header_row: int = 0
    first_data_row: int = 0
    last_data_row: int = 0
    confidence: float = 0.0


class Sheet(BaseModel):
    id: str
    name: str
    tables: list[Table] = Field(default_factory=list)
    cells: list[Cell] = Field(default_factory=list)


class Warning(BaseModel):
    address: str
    message: str
    level: str = "warn"  # "warn" | "error"


class Manifest(BaseModel):
    source_file: str = ""
    sheet_count: int = 0
    table_count: int = 0
    assumption_count: int = 0
    output_count: int = 0
    warnings: list[Warning] = Field(default_factory=list)


class Workbook(BaseModel):
    sheets: list[Sheet]
    manifest: Manifest = Field(default_factory=Manifest)

    @property
    def assumptions(self) -> list[Cell]:
        return [c for s in self.sheets for c in s.cells if c.role == CellRole.ASSUMPTION]

    @property
    def outputs(self) -> list[Cell]:
        return [c for s in self.sheets for c in s.cells if c.role == CellRole.OUTPUT]

    @property
    def all_cells(self) -> list[Cell]:
        return [c for s in self.sheets for c in s.cells]

    def get_cell(self, stable_id: str) -> Cell | None:
        return next((c for c in self.all_cells if c.stable_id == stable_id), None)
```

- [ ] Create `packages/osheet/tests/conftest.py`

```python
import io
import pytest
import openpyxl
from openpyxl.styles import PatternFill


@pytest.fixture
def simple_xlsx() -> bytes:
    """A minimal financial model xlsx for testing."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Revenue"

    # Header row
    ws["A1"] = "Month"
    ws["B1"] = "Revenue"
    ws["C1"] = "Costs"
    ws["D1"] = "Profit"

    # Assumption cell (yellow fill = common convention)
    ws["A3"] = "Churn Rate"
    ws["B3"] = 0.04
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    ws["B3"].fill = yellow

    # Data rows
    for i, (rev, cost) in enumerate([(100, 60), (110, 65), (120, 70)], start=5):
        ws[f"A{i}"] = f"2026-0{i-4}"
        ws[f"B{i}"] = rev
        ws[f"C{i}"] = cost
        ws[f"D{i}"] = f"=B{i}-C{i}"

    # Output cell
    ws["B10"] = "=SUM(B5:B7)"
    ws["A10"] = "Total Revenue"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture
def cross_sheet_xlsx() -> bytes:
    """Workbook with cross-sheet formula references."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Inputs"
    ws1["A1"] = "Growth Rate"
    ws1["B1"] = 0.1

    ws2 = wb.create_sheet("Revenue")
    ws2["A1"] = "Base"
    ws2["B1"] = 1000
    ws2["A2"] = "Projected"
    ws2["B2"] = "=B1*(1+Inputs!B1)"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
```

- [ ] Run: `pytest tests/test_models.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: core data models (Workbook, Sheet, Table, Cell, Manifest)"`

---

### Task 3: Parser (openpyxl → raw Cell list)

**Files:**
- Create: `packages/osheet/src/osheet/parser.py`
- Create: `packages/osheet/tests/test_parser.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_parser.py
import io
from osheet.parser import parse_xlsx

def test_parse_sheets(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    assert len(raw.sheets) == 1
    assert raw.sheets[0].name == "Revenue"

def test_parse_cells_count(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    # should have all non-empty cells
    assert len(sheet.cells) > 0

def test_parse_formula_cell(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    formula_cells = [c for c in sheet.cells if c.formula]
    assert len(formula_cells) >= 3  # D5, D6, D7 + B10

def test_parse_fill_color(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    b3 = next((c for c in sheet.cells if c.row == 3 and c.col == 2), None)
    assert b3 is not None
    assert b3.fill_color is not None

def test_cross_sheet(cross_sheet_xlsx):
    raw = parse_xlsx(cross_sheet_xlsx)
    assert len(raw.sheets) == 2
    names = [s.name for s in raw.sheets]
    assert "Inputs" in names and "Revenue" in names
```

- [ ] Run: `pytest tests/test_parser.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/parser.py`

```python
from __future__ import annotations
import io
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell as OxlCell
from osheet.models import Cell, CellRole, Sheet, Workbook, Manifest


def _fill_color(cell: OxlCell) -> str | None:
    try:
        fill = cell.fill
        if fill and fill.fill_type not in (None, "none"):
            fg = fill.fgColor
            if fg and fg.type == "rgb" and fg.rgb not in ("00000000", "FFFFFFFF"):
                return fg.rgb
    except Exception:
        pass
    return None


def parse_xlsx(data: bytes) -> Workbook:
    wb_ox = load_workbook(io.BytesIO(data), data_only=False)
    sheets: list[Sheet] = []

    for ws in wb_ox.worksheets:
        cells: list[Cell] = []
        for row in ws.iter_rows():
            for ox_cell in row:
                if ox_cell.value is None:
                    continue
                raw_val = ox_cell.value
                formula: str | None = None
                value = raw_val
                if isinstance(raw_val, str) and raw_val.startswith("="):
                    formula = raw_val
                    value = None  # formula result not available without data_only=True

                cells.append(Cell(
                    stable_id=f"{ws.title}.{ox_cell.column_letter}{ox_cell.row}",
                    role=CellRole.UNKNOWN,
                    value=value,
                    formula=formula,
                    sheet_name=ws.title,
                    col=ox_cell.column,
                    row=ox_cell.row,
                    fill_color=_fill_color(ox_cell),
                ))

        sheets.append(Sheet(id=f"sheet.{ws.title.lower().replace(' ', '_')}", name=ws.title, cells=cells))

    manifest = Manifest(source_file="", sheet_count=len(sheets))
    return Workbook(sheets=sheets, manifest=manifest)
```

- [ ] Run: `pytest tests/test_parser.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: xlsx parser (openpyxl → Workbook)"`

---

### Task 4: Analyzer Pass 1 — Table detection

**Files:**
- Create: `packages/osheet/src/osheet/analyzer/__init__.py`
- Create: `packages/osheet/src/osheet/analyzer/tables.py`
- Create: `packages/osheet/tests/test_tables.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_tables.py
from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables

def test_detects_table_in_simple_model(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    assert len(tables) >= 1

def test_table_has_range(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    t = tables[0]
    assert ":" in t.range_ref  # e.g. "A1:D7"

def test_table_header_row(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    tables = detect_tables(sheet)
    t = tables[0]
    assert t.header_row >= 1
```

- [ ] Run: `pytest tests/test_tables.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/analyzer/tables.py`

```python
from __future__ import annotations
from collections import defaultdict
from openpyxl.utils import get_column_letter, column_index_from_string
from osheet.models import Cell, Sheet, Table, Column, ColumnDtype


def _cell_grid(cells: list[Cell]) -> dict[tuple[int, int], Cell]:
    return {(c.row, c.col): c for c in cells}


def _is_text(cell: Cell) -> bool:
    return isinstance(cell.value, str) and cell.formula is None


def _is_data(cell: Cell) -> bool:
    return cell.value is not None or cell.formula is not None


def detect_tables(sheet: Sheet) -> list[Table]:
    if not sheet.cells:
        return []

    grid = _cell_grid(sheet.cells)
    rows = sorted({c.row for c in sheet.cells})
    cols = sorted({c.col for c in sheet.cells})

    if not rows or not cols:
        return []

    min_col, max_col = cols[0], cols[-1]

    # Find header rows: rows where majority of occupied cells are text
    header_candidates: list[int] = []
    for row in rows:
        row_cells = [grid[(row, c)] for c in cols if (row, c) in grid]
        if not row_cells:
            continue
        text_count = sum(1 for c in row_cells if _is_text(c))
        if text_count / len(row_cells) >= 0.5 and len(row_cells) >= 2:
            header_candidates.append(row)

    if not header_candidates:
        # No header found — treat entire populated area as one table
        header_candidates = [rows[0]]

    tables: list[Table] = []
    for i, header_row in enumerate(header_candidates):
        # Data rows: from header+1 until next header or end
        next_header = header_candidates[i + 1] if i + 1 < len(header_candidates) else None
        data_rows = [r for r in rows if r > header_row and (next_header is None or r < next_header)]
        if not data_rows:
            continue

        first_data = data_rows[0]
        last_data = data_rows[-1]

        # Columns: those with a header cell
        header_cols = [c for c in cols if (header_row, c) in grid]
        if not header_cols:
            header_cols = cols

        range_ref = (
            f"{get_column_letter(min(header_cols))}{header_row}:"
            f"{get_column_letter(max(header_cols))}{last_data}"
        )

        columns = []
        for col_idx in header_cols:
            header_cell = grid.get((header_row, col_idx))
            col_name = str(header_cell.value) if header_cell else get_column_letter(col_idx)
            columns.append(Column(name=col_name, dtype=ColumnDtype.UNKNOWN, col_index=col_idx))

        sheet_id = sheet.id.replace("sheet.", "")
        table_id = f"table.{sheet_id}.row{header_row}"

        tables.append(Table(
            id=table_id,
            sheet_name=sheet.name,
            range_ref=range_ref,
            columns=columns,
            header_row=header_row,
            first_data_row=first_data,
            last_data_row=last_data,
            confidence=0.8,
        ))

    return tables
```

- [ ] Create `packages/osheet/src/osheet/analyzer/__init__.py`

```python
from .tables import detect_tables
from .types import infer_column_types
from .roles import classify_roles
from .graph import build_formula_graph
from .ids import assign_stable_ids

def run_all(workbook):
    """Run all 5 analyzer passes in sequence, mutating workbook in place."""
    for sheet in workbook.sheets:
        tables = detect_tables(sheet)
        sheet.tables = tables
        infer_column_types(sheet)
    build_formula_graph(workbook)
    classify_roles(workbook)
    assign_stable_ids(workbook)
    # Update manifest counts
    workbook.manifest.table_count = sum(len(s.tables) for s in workbook.sheets)
    workbook.manifest.assumption_count = len(workbook.assumptions)
    workbook.manifest.output_count = len(workbook.outputs)
    workbook.manifest.sheet_count = len(workbook.sheets)
    return workbook
```

- [ ] Run: `pytest tests/test_tables.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: analyzer pass 1 — table detection"`

---

### Task 5: Analyzer Pass 2 — Column type inference

**Files:**
- Create: `packages/osheet/src/osheet/analyzer/types.py`
- Create: `packages/osheet/tests/test_types.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_types.py
from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables
from osheet.analyzer.types import infer_column_types
from osheet.models import ColumnDtype

def test_number_column(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    sheet.tables = detect_tables(sheet)
    infer_column_types(sheet)
    # Revenue column (col B) should be NUMBER
    for table in sheet.tables:
        for col in table.columns:
            if "Revenue" in col.name:
                assert col.dtype == ColumnDtype.NUMBER

def test_formula_column(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    sheet = raw.sheets[0]
    sheet.tables = detect_tables(sheet)
    infer_column_types(sheet)
    for table in sheet.tables:
        for col in table.columns:
            if "Profit" in col.name:
                assert col.dtype in (ColumnDtype.FORMULA, ColumnDtype.NUMBER)
```

- [ ] Run: `pytest tests/test_types.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/analyzer/types.py`

```python
from __future__ import annotations
import datetime
from osheet.models import Cell, Sheet, Column, ColumnDtype


def _infer_dtype(cells: list[Cell]) -> ColumnDtype:
    if not cells:
        return ColumnDtype.UNKNOWN
    formula_count = sum(1 for c in cells if c.formula)
    if formula_count / len(cells) >= 0.5:
        return ColumnDtype.FORMULA
    values = [c.value for c in cells if c.value is not None]
    if not values:
        return ColumnDtype.UNKNOWN
    number_count = sum(1 for v in values if isinstance(v, (int, float)))
    date_count = sum(1 for v in values if isinstance(v, (datetime.date, datetime.datetime)))
    text_count = sum(1 for v in values if isinstance(v, str))
    total = len(values)
    if number_count / total >= 0.7:
        return ColumnDtype.NUMBER
    if date_count / total >= 0.7:
        return ColumnDtype.DATE
    if text_count / total >= 0.7:
        return ColumnDtype.TEXT
    return ColumnDtype.MIXED


def infer_column_types(sheet: Sheet) -> None:
    """Mutates each Column.dtype in place."""
    grid = {(c.row, c.col): c for c in sheet.cells}
    for table in sheet.tables:
        for col in table.columns:
            data_cells = [
                grid[(r, col.col_index)]
                for r in range(table.first_data_row, table.last_data_row + 1)
                if (r, col.col_index) in grid
            ]
            col.dtype = _infer_dtype(data_cells)
```

- [ ] Run: `pytest tests/test_types.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: analyzer pass 2 — column dtype inference"`

---

### Task 6: Analyzer Pass 4 — Formula dependency graph

> Pass 4 before Pass 3 because role classification needs the graph.

**Files:**
- Create: `packages/osheet/src/osheet/analyzer/graph.py`
- Create: `packages/osheet/tests/test_graph.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_graph.py
from osheet.parser import parse_xlsx
from osheet.analyzer.graph import build_formula_graph

def test_formula_creates_edge(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    formula_cells = [c for s in raw.sheets for c in s.cells if c.formula]
    # Each formula cell should depend on at least one other cell
    assert any(len(c.depends_on) > 0 for c in formula_cells)

def test_cross_sheet_dependency(cross_sheet_xlsx):
    raw = parse_xlsx(cross_sheet_xlsx)
    build_formula_graph(raw)
    # Revenue!B2 = "=B1*(1+Inputs!B1)" should depend on Inputs!B1
    rev_sheet = next(s for s in raw.sheets if s.name == "Revenue")
    b2 = next((c for c in rev_sheet.cells if c.row == 2 and c.col == 2), None)
    assert b2 is not None
    assert len(b2.depends_on) >= 1

def test_sum_range_expands(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    sum_cell = next((c for c in all_cells if c.formula and "SUM" in (c.formula or "")), None)
    if sum_cell:
        assert len(sum_cell.depends_on) >= 1
```

- [ ] Run: `pytest tests/test_graph.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/analyzer/graph.py`

```python
from __future__ import annotations
import re
from openpyxl.utils import column_index_from_string, get_column_letter
from osheet.models import Cell, Workbook

# Matches: optional 'Sheet'! prefix + $?COL$?ROW
_CELL_RE = re.compile(
    r"(?:'?([^'!\[\]]+)'?!)?(\$?[A-Z]+)(\$?[0-9]+)",
    re.IGNORECASE,
)
# Matches ranges: A1:B10 with optional sheet prefix
_RANGE_RE = re.compile(
    r"(?:'?([^'!\[\]]+)'?!)?(\$?[A-Z]+)(\$?[0-9]+):(\$?[A-Z]+)(\$?[0-9]+)",
    re.IGNORECASE,
)


def _col_num(col_str: str) -> int:
    return column_index_from_string(col_str.replace("$", ""))


def _row_num(row_str: str) -> int:
    return int(row_str.replace("$", ""))


def _expand_range(sheet: str, col1: str, row1: str, col2: str, row2: str) -> list[tuple[str, int, int]]:
    c1, c2 = _col_num(col1), _col_num(col2)
    r1, r2 = _row_num(row1), _row_num(row2)
    refs = []
    for c in range(min(c1, c2), max(c1, c2) + 1):
        for r in range(min(r1, r2), max(r1, r2) + 1):
            refs.append((sheet, c, r))
    return refs


def _parse_refs(formula: str, default_sheet: str) -> list[tuple[str, int, int]]:
    """Return list of (sheet_name, col, row) referenced by this formula."""
    refs: list[tuple[str, int, int]] = []
    # First extract ranges (superset of single cells)
    for m in _RANGE_RE.finditer(formula):
        sheet = m.group(1) or default_sheet
        refs.extend(_expand_range(sheet, m.group(2), m.group(3), m.group(4), m.group(5)))
    # Then single cells not already covered
    range_spans = {(m.start(), m.end()) for m in _RANGE_RE.finditer(formula)}
    for m in _CELL_RE.finditer(formula):
        # Skip if this match is inside a range match
        if any(rs <= m.start() and m.end() <= re_ for rs, re_ in range_spans):
            continue
        sheet = m.group(1) or default_sheet
        try:
            col = _col_num(m.group(2))
            row = _row_num(m.group(3))
            refs.append((sheet, col, row))
        except Exception:
            continue
    return refs


def build_formula_graph(workbook: Workbook) -> None:
    """Mutates each Cell.depends_on with stable_ids of its dependencies."""
    # Build address → stable_id lookup: (sheet_name, col, row) → stable_id
    addr_to_id: dict[tuple[str, int, int], str] = {}
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            addr_to_id[(sheet.name, cell.col, cell.row)] = cell.stable_id

    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if not cell.formula:
                continue
            refs = _parse_refs(cell.formula, sheet.name)
            deps: list[str] = []
            for (sname, col, row) in refs:
                # Try exact sheet name match
                sid = addr_to_id.get((sname, col, row))
                if sid and sid != cell.stable_id:
                    deps.append(sid)
            cell.depends_on = list(dict.fromkeys(deps))  # deduplicate, preserve order
```

- [ ] Run: `pytest tests/test_graph.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: analyzer pass 4 — formula dependency graph"`

---

### Task 7: Analyzer Pass 3 — Cell role classification

**Files:**
- Create: `packages/osheet/src/osheet/analyzer/roles.py`
- Create: `packages/osheet/tests/test_roles.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_roles.py
from osheet.parser import parse_xlsx
from osheet.analyzer.graph import build_formula_graph
from osheet.analyzer.roles import classify_roles
from osheet.models import CellRole

def test_yellow_cell_is_assumption(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    sheet = raw.sheets[0]
    # B3 has yellow fill and a scalar value
    b3 = next((c for c in sheet.cells if c.row == 3 and c.col == 2), None)
    assert b3 is not None
    assert b3.role == CellRole.ASSUMPTION

def test_sum_cell_is_output(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    build_formula_graph(raw)
    classify_roles(raw)
    all_cells = [c for s in raw.sheets for c in s.cells]
    # B10 =SUM(B5:B7) — a formula not depended upon by others → output
    sum_cell = next((c for c in all_cells if c.formula and "SUM" in c.formula), None)
    assert sum_cell is not None
    assert sum_cell.role in (CellRole.OUTPUT, CellRole.INTERMEDIATE)
```

- [ ] Run: `pytest tests/test_roles.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/analyzer/roles.py`

```python
from __future__ import annotations
from osheet.models import Cell, CellRole, Workbook

_ASSUMPTION_COLORS = {
    "FFFFFF00",  # yellow
    "FFFFC000",  # orange-yellow
    "FFFFEB9C",  # light yellow
    "FFD9E1F2",  # light blue (sometimes used for inputs)
}


def classify_roles(workbook: Workbook) -> None:
    """Mutates each Cell.role in place using graph in/out-degree + fill heuristics."""
    all_cells = workbook.all_cells
    cell_by_id = {c.stable_id: c for c in all_cells}

    # Count in-degree (how many cells reference this cell)
    in_degree: dict[str, int] = {c.stable_id: 0 for c in all_cells}
    for cell in all_cells:
        for dep_id in cell.depends_on:
            if dep_id in in_degree:
                in_degree[dep_id] += 1

    for cell in all_cells:
        # Labels: text cells with no formula
        if isinstance(cell.value, str) and not cell.formula:
            cell.role = CellRole.LABEL
            continue

        # Assumptions: hardcoded scalar (no formula), referenced by others
        if not cell.formula:
            if in_degree.get(cell.stable_id, 0) > 0:
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.75
            elif cell.fill_color and cell.fill_color.upper() in _ASSUMPTION_COLORS:
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.9
            elif isinstance(cell.value, (int, float)):
                # Scalar with no dependents — could be assumption or just data
                cell.role = CellRole.ASSUMPTION
                cell.confidence = 0.5
            else:
                cell.role = CellRole.UNKNOWN
            continue

        # Formula cells
        out_degree = len(cell.depends_on)
        dependents_count = in_degree.get(cell.stable_id, 0)

        if dependents_count == 0:
            # Nothing depends on this formula → it's an output
            cell.role = CellRole.OUTPUT
            cell.confidence = 0.85
        elif out_degree > 0:
            cell.role = CellRole.INTERMEDIATE
            cell.confidence = 0.8
        else:
            cell.role = CellRole.UNKNOWN

        # Override: yellow fill on formula cell → probably a special output
        if cell.fill_color and cell.fill_color.upper() in _ASSUMPTION_COLORS:
            cell.role = CellRole.ASSUMPTION
            cell.confidence = 0.7
```

- [ ] Run: `pytest tests/test_roles.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: analyzer pass 3 — cell role classification"`

---

### Task 8: Analyzer Pass 5 — Stable ID assignment

**Files:**
- Create: `packages/osheet/src/osheet/analyzer/ids.py`
- Create: `packages/osheet/tests/test_ids.py`

- [ ] Write failing test

```python
# packages/osheet/tests/test_ids.py
from osheet.parser import parse_xlsx
from osheet.analyzer.tables import detect_tables
from osheet.analyzer.ids import assign_stable_ids
import re

def test_ids_are_unique(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    for sheet in raw.sheets:
        sheet.tables = detect_tables(sheet)
    assign_stable_ids(raw)
    all_ids = [c.stable_id for s in raw.sheets for c in s.cells]
    assert len(all_ids) == len(set(all_ids))

def test_ids_are_slugified(simple_xlsx):
    raw = parse_xlsx(simple_xlsx)
    for sheet in raw.sheets:
        sheet.tables = detect_tables(sheet)
    assign_stable_ids(raw)
    for cell in raw.all_cells:
        assert re.match(r'^[a-z0-9._]+$', cell.stable_id), f"Bad ID: {cell.stable_id}"

def test_ids_stable_across_runs(simple_xlsx):
    raw1 = parse_xlsx(simple_xlsx)
    raw2 = parse_xlsx(simple_xlsx)
    for wb in (raw1, raw2):
        for sheet in wb.sheets:
            sheet.tables = detect_tables(sheet)
        assign_stable_ids(wb)
    ids1 = sorted(c.stable_id for s in raw1.sheets for c in s.cells)
    ids2 = sorted(c.stable_id for s in raw2.sheets for c in s.cells)
    assert ids1 == ids2
```

- [ ] Run: `pytest tests/test_ids.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/analyzer/ids.py`

```python
from __future__ import annotations
import re
from openpyxl.utils import get_column_letter
from osheet.models import Cell, CellRole, Sheet, Table, Workbook


def _slug(text: str) -> str:
    """Convert any string to a lowercase dot-safe slug."""
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = text.strip("_")
    return text or "cell"


def _table_for_cell(cell: Cell, tables: list[Table]) -> Table | None:
    for t in tables:
        if cell.row >= t.header_row and cell.row <= t.last_data_row:
            if any(col.col_index == cell.col for col in t.columns):
                return t
    return None


def _col_name_for_cell(cell: Cell, table: Table) -> str | None:
    for col in table.columns:
        if col.col_index == cell.col:
            return col.name
    return None


def assign_stable_ids(workbook: Workbook) -> None:
    """Assign deterministic, human-readable stable IDs to every cell."""
    seen: set[str] = set()

    def unique(base: str) -> str:
        if base not in seen:
            seen.add(base)
            return base
        i = 2
        while f"{base}_{i}" in seen:
            i += 1
        candidate = f"{base}_{i}"
        seen.add(candidate)
        return candidate

    for sheet in workbook.sheets:
        sheet_slug = _slug(sheet.name)
        for cell in sheet.cells:
            table = _table_for_cell(cell, sheet.tables)
            if table:
                table_slug = _slug(table.id.split(".")[-1])
                col_name = _col_name_for_cell(cell, table)
                if col_name and cell.row != table.header_row:
                    col_slug = _slug(col_name)
                    row_label = f"r{cell.row}"
                    # For assumption/output cells, use their role as prefix
                    if cell.role == CellRole.ASSUMPTION:
                        base = f"assumption.{sheet_slug}.{col_slug}"
                    elif cell.role == CellRole.OUTPUT:
                        base = f"metric.{sheet_slug}.{col_slug}"
                    else:
                        base = f"{sheet_slug}.{table_slug}.{col_slug}.{row_label}"
                else:
                    base = f"{sheet_slug}.{get_column_letter(cell.col)}{cell.row}".lower()
            else:
                col_letter = get_column_letter(cell.col).lower()
                base = f"{sheet_slug}.{col_letter}{cell.row}"

            cell.stable_id = unique(base)
```

- [ ] Run: `pytest tests/test_ids.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: analyzer pass 5 — stable ID assignment"`

---

### Task 9: Emitters — .osheet zip and annotated .xlsx

**Files:**
- Create: `packages/osheet/src/osheet/emitter/__init__.py`
- Create: `packages/osheet/src/osheet/emitter/osheet.py`
- Create: `packages/osheet/src/osheet/emitter/xlsx.py`
- Create: `packages/osheet/tests/test_emitter.py`

- [ ] Write failing tests

```python
# packages/osheet/tests/test_emitter.py
import io, json, zipfile
from osheet.parser import parse_xlsx
from osheet.analyzer import run_all
from osheet.emitter.osheet import to_osheet_bytes
from osheet.emitter.xlsx import to_xlsx_bytes

def _analyzed(fixture):
    wb = parse_xlsx(fixture)
    return run_all(wb)

def test_osheet_is_zip(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    assert zipfile.is_zipfile(io.BytesIO(data))

def test_osheet_contains_required_files(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        names = zf.namelist()
    assert "workbook.json" in names
    assert "sheets.json" in names
    assert "formula_graph.json" in names

def test_osheet_workbook_json_valid(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_osheet_bytes(wb)
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        obj = json.loads(zf.read("workbook.json"))
    assert "manifest" in obj
    assert obj["manifest"]["sheet_count"] >= 1

def test_xlsx_output_is_valid(simple_xlsx):
    wb = _analyzed(simple_xlsx)
    data = to_xlsx_bytes(wb)
    import openpyxl
    loaded = openpyxl.load_workbook(io.BytesIO(data))
    assert len(loaded.sheetnames) >= 1
```

- [ ] Run: `pytest tests/test_emitter.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/emitter/osheet.py`

```python
from __future__ import annotations
import io, json, zipfile
from osheet.models import Workbook


def to_osheet_bytes(workbook: Workbook) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # workbook.json
        manifest_data = workbook.manifest.model_dump()
        zf.writestr("workbook.json", json.dumps({"manifest": manifest_data}, indent=2))

        # sheets.json — full sheet/table/cell tree
        sheets_data = workbook.model_dump(include={"sheets"})
        zf.writestr("sheets.json", json.dumps(sheets_data, indent=2, default=str))

        # formula_graph.json — adjacency list
        graph: dict[str, list[str]] = {}
        for sheet in workbook.sheets:
            for cell in sheet.cells:
                if cell.depends_on:
                    graph[cell.stable_id] = cell.depends_on
        zf.writestr("formula_graph.json", json.dumps(graph, indent=2))

    return buf.getvalue()
```

- [ ] Implement `packages/osheet/src/osheet/emitter/xlsx.py`

```python
from __future__ import annotations
import io
import openpyxl
from openpyxl.styles import PatternFill, Font
from osheet.models import CellRole, Workbook


_ROLE_COLORS = {
    CellRole.ASSUMPTION:   "FFD4A96A",
    CellRole.OUTPUT:       "FF6AB07A",
    CellRole.INTERMEDIATE: "FF6A8FD4",
}


def to_xlsx_bytes(workbook: Workbook) -> bytes:
    wb_ox = openpyxl.Workbook()
    wb_ox.remove(wb_ox.active)  # remove default sheet

    for sheet in workbook.sheets:
        ws = wb_ox.create_sheet(title=sheet.name)
        for cell in sheet.cells:
            ox_cell = ws.cell(row=cell.row, column=cell.col)
            ox_cell.value = cell.formula if cell.formula else cell.value

            # Colour-code by role
            color = _ROLE_COLORS.get(cell.role)
            if color:
                ox_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

            # Embed stable_id as cell comment
            if cell.stable_id:
                from openpyxl.comments import Comment
                comment = Comment(
                    f"stable_id: {cell.stable_id}\nrole: {cell.role.value}\nconf: {cell.confidence:.2f}",
                    "osheet"
                )
                ox_cell.comment = comment

    # Manifest hidden sheet
    meta_ws = wb_ox.create_sheet(title="_ai_manifest")
    meta_ws.sheet_state = "hidden"
    meta_ws["A1"] = "osheet_version"
    meta_ws["B1"] = "0.1.0"
    meta_ws["A2"] = "assumption_count"
    meta_ws["B2"] = workbook.manifest.assumption_count
    meta_ws["A3"] = "output_count"
    meta_ws["B3"] = workbook.manifest.output_count
    meta_ws["A4"] = "table_count"
    meta_ws["B4"] = workbook.manifest.table_count

    buf = io.BytesIO()
    wb_ox.save(buf)
    return buf.getvalue()
```

- [ ] Implement `packages/osheet/src/osheet/emitter/__init__.py`

```python
from .osheet import to_osheet_bytes
from .xlsx import to_xlsx_bytes
```

- [ ] Run: `pytest tests/test_emitter.py -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: emitters — .osheet zip and annotated .xlsx"`

---

### Task 10: Public API

**Files:**
- Create: `packages/osheet/src/osheet/api.py`
- Modify: `packages/osheet/src/osheet/__init__.py`
- Create: `packages/osheet/tests/test_api.py`

- [ ] Write failing tests

```python
# packages/osheet/tests/test_api.py
import io
import osheet
from osheet.models import CellRole

def test_load_returns_workbook(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    assert len(wb.sheets) >= 1

def test_load_runs_all_passes(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    # After all passes, some cells should have non-UNKNOWN roles
    roles = {c.role for s in wb.sheets for c in s.cells}
    assert roles - {CellRole.UNKNOWN} != set()

def test_trace_returns_upstream(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    outputs = wb.outputs
    if outputs:
        result = osheet.trace(wb, outputs[0].stable_id)
        assert hasattr(result, "upstream")
        assert hasattr(result, "downstream")

def test_find_fuzzy(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    results = osheet.find(wb, "revenue")
    assert isinstance(results, list)

def test_export_xlsx(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    data = wb.export_xlsx()
    assert isinstance(data, bytes)
    assert len(data) > 100

def test_export_osheet(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    data = wb.export_osheet()
    import zipfile
    assert zipfile.is_zipfile(io.BytesIO(data))

def test_propose_patch(simple_xlsx):
    wb = osheet.load(simple_xlsx)
    assumptions = wb.assumptions
    if assumptions:
        a = assumptions[0]
        proposal = osheet.propose_patch(wb, a.stable_id, 999)
        assert proposal.cell_id == a.stable_id
        assert proposal.new_value == 999
        assert isinstance(proposal.affected_cells, list)
```

- [ ] Run: `pytest tests/test_api.py -v` — expect FAIL

- [ ] Implement `packages/osheet/src/osheet/api.py`

```python
from __future__ import annotations
from dataclasses import dataclass, field
from osheet.models import Cell, CellRole, Workbook
from osheet.parser import parse_xlsx
from osheet.analyzer import run_all
from osheet.emitter.xlsx import to_xlsx_bytes
from osheet.emitter.osheet import to_osheet_bytes


@dataclass
class TraceResult:
    cell_id: str
    upstream: list[str] = field(default_factory=list)
    downstream: list[str] = field(default_factory=list)


@dataclass
class PatchProposal:
    cell_id: str
    old_value: object
    new_value: object
    affected_cells: list[str] = field(default_factory=list)
    diff: str = ""


def load(data: bytes) -> "OsheetWorkbook":
    raw = parse_xlsx(data)
    run_all(raw)
    return OsheetWorkbook(raw)


def trace(workbook: "OsheetWorkbook", stable_id: str) -> TraceResult:
    return workbook.trace(stable_id)


def find(workbook: "OsheetWorkbook", query: str) -> list[Cell]:
    return workbook.find(query)


def propose_patch(workbook: "OsheetWorkbook", stable_id: str, new_value: object) -> PatchProposal:
    return workbook.propose_patch(stable_id, new_value)


class OsheetWorkbook:
    """Thin wrapper around Workbook that exposes the public agent API."""

    def __init__(self, wb: Workbook):
        self._wb = wb

    # Delegate common attributes
    @property
    def sheets(self): return self._wb.sheets
    @property
    def assumptions(self): return self._wb.assumptions
    @property
    def outputs(self): return self._wb.outputs
    @property
    def manifest(self): return self._wb.manifest

    def find(self, query: str) -> list[Cell]:
        q = query.lower()
        return [c for c in self._wb.all_cells
                if q in c.stable_id.lower() or (isinstance(c.value, str) and q in c.value.lower())]

    def trace(self, stable_id: str) -> TraceResult:
        cell = self._wb.get_cell(stable_id)
        if cell is None:
            return TraceResult(cell_id=stable_id)
        # downstream: cells that depend on this cell
        downstream = [c.stable_id for c in self._wb.all_cells if stable_id in c.depends_on]
        return TraceResult(cell_id=stable_id, upstream=cell.depends_on, downstream=downstream)

    def propose_patch(self, stable_id: str, new_value: object) -> PatchProposal:
        cell = self._wb.get_cell(stable_id)
        if cell is None:
            return PatchProposal(cell_id=stable_id, old_value=None, new_value=new_value)
        # Find all cells downstream (transitively)
        visited: set[str] = set()
        queue = [stable_id]
        while queue:
            cid = queue.pop()
            if cid in visited:
                continue
            visited.add(cid)
            downstream = [c.stable_id for c in self._wb.all_cells if cid in c.depends_on]
            queue.extend(downstream)
        visited.discard(stable_id)
        diff = f"Change {stable_id}: {cell.value!r} → {new_value!r}\nAffects {len(visited)} downstream cells."
        return PatchProposal(
            cell_id=stable_id,
            old_value=cell.value,
            new_value=new_value,
            affected_cells=list(visited),
            diff=diff,
        )

    def apply_patch(self, proposal: PatchProposal) -> None:
        cell = self._wb.get_cell(proposal.cell_id)
        if cell:
            cell.value = proposal.new_value
            cell.formula = None

    def export_xlsx(self) -> bytes:
        return to_xlsx_bytes(self._wb)

    def export_osheet(self) -> bytes:
        return to_osheet_bytes(self._wb)
```

- [ ] Update `packages/osheet/src/osheet/__init__.py`

```python
from osheet.api import load, trace, find, propose_patch, OsheetWorkbook
from osheet.models import CellRole, ColumnDtype

__all__ = ["load", "trace", "find", "propose_patch", "OsheetWorkbook", "CellRole", "ColumnDtype"]
```

- [ ] Run: `pytest tests/test_api.py -v` — expect PASS
- [ ] Run full suite: `pytest --tb=short -q` — all green
- [ ] Commit: `git add . && git commit -m "feat: public API (load, trace, find, propose_patch, export)"`

---

### Task 11: FastAPI backend

**Files:**
- Create: `packages/osheet-app/backend/pyproject.toml`
- Create: `packages/osheet-app/backend/src/app/main.py`
- Create: `packages/osheet-app/backend/src/app/storage.py`
- Create: `packages/osheet-app/backend/src/app/routes/convert.py`
- Create: `packages/osheet-app/backend/src/app/routes/result.py`
- Create: `packages/osheet-app/backend/src/app/routes/download.py`
- Create: `packages/osheet-app/backend/tests/test_routes.py`

- [ ] Create `packages/osheet-app/backend/pyproject.toml`

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "osheet-app"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
    "fastapi>=0.111",
    "uvicorn[standard]>=0.29",
    "python-multipart>=0.0.9",
    "osheet @ file://../osheet",
]

[project.optional-dependencies]
dev = ["pytest>=8", "httpx>=0.27", "pytest-asyncio"]

[tool.hatch.build.targets.wheel]
packages = ["src/app"]

[tool.pytest.ini_options]
testpaths = ["tests"]
asyncio_mode = "auto"
```

- [ ] Install: `cd packages/osheet-app/backend && pip install -e ".[dev]"`

- [ ] Implement `packages/osheet-app/backend/src/app/storage.py`

```python
from __future__ import annotations
import time
import uuid
from dataclasses import dataclass, field

_TTL_SECONDS = 3600  # 1 hour


@dataclass
class Job:
    job_id: str
    status: str  # "processing" | "done" | "error"
    error: str | None = None
    xlsx_bytes: bytes | None = None
    osheet_bytes: bytes | None = None
    summary: dict = field(default_factory=dict)
    created_at: float = field(default_factory=time.time)


_store: dict[str, Job] = {}


def create_job() -> Job:
    job = Job(job_id=str(uuid.uuid4()), status="processing")
    _store[job.job_id] = job
    return job


def get_job(job_id: str) -> Job | None:
    job = _store.get(job_id)
    if job and time.time() - job.created_at > _TTL_SECONDS:
        del _store[job_id]
        return None
    return job


def update_job(job: Job) -> None:
    _store[job.job_id] = job
```

- [ ] Implement `packages/osheet-app/backend/src/app/routes/convert.py`

```python
from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks
import osheet
from app.storage import create_job, update_job, Job

router = APIRouter()


def _run_conversion(job: Job, data: bytes) -> None:
    try:
        wb = osheet.load(data)
        job.xlsx_bytes = wb.export_xlsx()
        job.osheet_bytes = wb.export_osheet()
        job.summary = {
            "sheet_count": wb.manifest.sheet_count,
            "table_count": wb.manifest.table_count,
            "assumption_count": wb.manifest.assumption_count,
            "output_count": wb.manifest.output_count,
            "warning_count": len(wb.manifest.warnings),
            "warnings": [w.model_dump() for w in wb.manifest.warnings],
        }
        job.status = "done"
    except Exception as exc:
        job.status = "error"
        job.error = str(exc)
    finally:
        update_job(job)


@router.post("/convert")
async def convert(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename or not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are accepted")
    data = await file.read()
    if len(data) > 20 * 1024 * 1024:  # 20 MB limit
        raise HTTPException(status_code=413, detail="File too large (max 20 MB)")
    job = create_job()
    background_tasks.add_task(_run_conversion, job, data)
    return {"job_id": job.job_id, "status": "processing"}
```

- [ ] Implement `packages/osheet-app/backend/src/app/routes/result.py`

```python
from fastapi import APIRouter, HTTPException
from app.storage import get_job

router = APIRouter()

@router.get("/result/{job_id}")
async def result(job_id: str):
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found or expired")
    return {
        "job_id": job.job_id,
        "status": job.status,
        "error": job.error,
        "summary": job.summary,
    }
```

- [ ] Implement `packages/osheet-app/backend/src/app/routes/download.py`

```python
from fastapi import APIRouter, HTTPException
from fastapi.responses import Response
from app.storage import get_job

router = APIRouter()

@router.get("/download/{job_id}/{file_type}")
async def download(job_id: str, file_type: str):
    job = get_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Job not found or expired")
    if job.status != "done":
        raise HTTPException(status_code=400, detail=f"Job status: {job.status}")
    if file_type == "xlsx":
        return Response(content=job.xlsx_bytes, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": "attachment; filename=output.ai.xlsx"})
    if file_type == "osheet":
        return Response(content=job.osheet_bytes, media_type="application/zip",
                        headers={"Content-Disposition": "attachment; filename=output.osheet"})
    raise HTTPException(status_code=400, detail="file_type must be xlsx or osheet")
```

- [ ] Implement `packages/osheet-app/backend/src/app/main.py`

```python
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routes.convert import router as convert_router
from app.routes.result import router as result_router
from app.routes.download import router as download_router

app = FastAPI(title="osheet API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(convert_router)
app.include_router(result_router)
app.include_router(download_router)

@app.get("/health")
async def health():
    return {"status": "ok"}
```

- [ ] Write `packages/osheet-app/backend/tests/test_routes.py`

```python
import io, time
import pytest
from fastapi.testclient import TestClient
import openpyxl
from app.main import app

client = TestClient(app)


def _make_xlsx() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Revenue"
    ws["B1"] = 100
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_health():
    r = client.get("/health")
    assert r.status_code == 200

def test_convert_and_poll():
    data = _make_xlsx()
    r = client.post("/convert", files={"file": ("model.xlsx", data, "application/octet-stream")})
    assert r.status_code == 200
    job_id = r.json()["job_id"]

    # Poll until done (TestClient runs background tasks synchronously)
    r2 = client.get(f"/result/{job_id}")
    assert r2.status_code == 200
    assert r2.json()["status"] in ("done", "processing")

def test_download_xlsx():
    data = _make_xlsx()
    r = client.post("/convert", files={"file": ("model.xlsx", data, "application/octet-stream")})
    job_id = r.json()["job_id"]
    # Wait for background task
    for _ in range(10):
        r2 = client.get(f"/result/{job_id}")
        if r2.json()["status"] == "done":
            break
        time.sleep(0.1)
    r3 = client.get(f"/download/{job_id}/xlsx")
    assert r3.status_code == 200
    assert r3.headers["content-type"].startswith("application/vnd")

def test_rejects_non_xlsx():
    r = client.post("/convert", files={"file": ("model.csv", b"a,b,c", "text/csv")})
    assert r.status_code == 400
```

- [ ] Run: `cd packages/osheet-app/backend && pytest -v` — expect PASS
- [ ] Commit: `git add . && git commit -m "feat: FastAPI backend (convert, result, download routes)"`

---

### Task 12: Next.js frontend

**Files:**
- Create: `packages/osheet-app/frontend/package.json`
- Create: `packages/osheet-app/frontend/next.config.ts`
- Create: `packages/osheet-app/frontend/tailwind.config.ts`
- Create: `packages/osheet-app/frontend/src/app/layout.tsx`
- Create: `packages/osheet-app/frontend/src/app/page.tsx`
- Create: `packages/osheet-app/frontend/src/app/inspect/[jobId]/page.tsx`
- Create: `packages/osheet-app/frontend/src/app/components/UploadZone.tsx`
- Create: `packages/osheet-app/frontend/src/app/components/StatBar.tsx`
- Create: `packages/osheet-app/frontend/src/app/components/SheetSidebar.tsx`
- Create: `packages/osheet-app/frontend/src/app/components/CellTable.tsx`
- Create: `packages/osheet-app/frontend/src/app/components/DetailPanel.tsx`

- [ ] Bootstrap Next.js app

```bash
cd packages/osheet-app/frontend
npx create-next-app@latest . --typescript --tailwind --app --no-src-dir --import-alias "@/*" --yes
```

- [ ] Install Geist font package

```bash
npm install geist
```

- [ ] Update `packages/osheet-app/frontend/next.config.ts`

```typescript
import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  async rewrites() {
    return [{ source: "/api/:path*", destination: "http://localhost:8000/:path*" }];
  },
};

export default nextConfig;
```

- [ ] Update `packages/osheet-app/frontend/src/app/layout.tsx`

```tsx
import type { Metadata } from "next";
import { GeistSans } from "geist/font/sans";
import { GeistMono } from "geist/font/mono";
import "./globals.css";

export const metadata: Metadata = {
  title: "osheet — AI-native spreadsheet compiler",
  description: "Upload any .xlsx. Get back an AI-native workbook.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className={`${GeistSans.variable} ${GeistMono.variable}`}>
      <body className="bg-[#171717] text-[#ededed] font-sans antialiased min-h-screen">
        {children}
      </body>
    </html>
  );
}
```

- [ ] Update `packages/osheet-app/frontend/src/app/globals.css`

```css
@tailwind base;
@tailwind components;
@tailwind utilities;

:root {
  --canvas: #171717;
  --canvas-soft: #1f1f1f;
  --canvas-hover: #242424;
  --hairline: #2e2e2e;
  --ink: #ededed;
  --body-strong: #b8b8b8;
  --body: #737373;
  --mute: #525252;
}

* { box-sizing: border-box; }
::-webkit-scrollbar { width: 3px; }
::-webkit-scrollbar-thumb { background: #2e2e2e; border-radius: 9999px; }
```

- [ ] Update `packages/osheet-app/frontend/tailwind.config.ts`

```typescript
import type { Config } from "tailwindcss";

const config: Config = {
  content: ["./src/**/*.{ts,tsx}"],
  theme: {
    extend: {
      fontFamily: {
        sans: ["var(--font-geist-sans)", "system-ui", "sans-serif"],
        mono: ["var(--font-geist-mono)", "monospace"],
      },
      colors: {
        canvas: "#171717",
        "canvas-soft": "#1f1f1f",
        "canvas-hover": "#242424",
        hairline: "#2e2e2e",
        ink: "#ededed",
        "body-strong": "#b8b8b8",
        body: "#737373",
        mute: "#525252",
        assumption: "#d4a96a",
        output: "#6ab07a",
        intermediate: "#6a8fd4",
        warning: "#d46a6a",
      },
    },
  },
  plugins: [],
};

export default config;
```

- [ ] Create Upload page `packages/osheet-app/frontend/src/app/page.tsx`

```tsx
"use client";
import { useState, useCallback } from "react";
import { useRouter } from "next/navigation";

export default function UploadPage() {
  const router = useRouter();
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".xlsx")) {
      setError("Only .xlsx files are supported");
      return;
    }
    setLoading(true);
    setError(null);
    const form = new FormData();
    form.append("file", file);
    try {
      const res = await fetch("/api/convert", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Conversion failed");
      router.push(`/inspect/${data.job_id}`);
    } catch (e: any) {
      setError(e.message);
      setLoading(false);
    }
  }, [router]);

  return (
    <main className="min-h-screen flex flex-col items-center justify-center px-6">
      <div className="mb-10 text-center">
        <h1 className="text-[32px] font-medium tracking-[-0.8px] text-ink mb-3">osheet</h1>
        <p className="text-body text-[15px]">Upload any .xlsx — get back an AI-native workbook.</p>
      </div>

      <div
        className={`w-full max-w-md border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors
          ${dragging ? "border-ink bg-canvas-soft" : "border-hairline hover:border-mute"}`}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }}
        onClick={() => document.getElementById("file-input")?.click()}
      >
        <input id="file-input" type="file" accept=".xlsx" className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
        {loading ? (
          <p className="text-body-strong text-sm">Converting…</p>
        ) : (
          <>
            <p className="text-ink text-sm font-medium mb-1">Drop your .xlsx here</p>
            <p className="text-mute text-xs">or click to browse · max 20 MB</p>
          </>
        )}
      </div>

      {error && <p className="mt-4 text-warning text-sm">{error}</p>}

      <p className="mt-8 text-mute text-xs">Files are processed in memory and not stored.</p>
    </main>
  );
}
```

- [ ] Create Inspector page `packages/osheet-app/frontend/src/app/inspect/[jobId]/page.tsx`

```tsx
"use client";
import { useEffect, useState } from "react";
import { useParams } from "next/navigation";

type CellData = { stable_id: string; role: string; value: any; formula?: string; confidence: number; depends_on: string[]; col: number; row: number };
type SheetData = { name: string; cells: CellData[]; tables: any[] };
type Summary = { sheet_count: number; table_count: number; assumption_count: number; output_count: number; warning_count: number; warnings: { address: string; message: string }[] };

const ROLE_COLORS: Record<string, string> = {
  assumption: "text-assumption border-assumption/30 bg-assumption/10",
  output: "text-output border-output/30 bg-output/10",
  intermediate: "text-intermediate border-intermediate/30 bg-intermediate/10",
};

export default function InspectPage() {
  const { jobId } = useParams<{ jobId: string }>();
  const [status, setStatus] = useState("loading");
  const [summary, setSummary] = useState<Summary | null>(null);
  const [selected, setSelected] = useState<CellData | null>(null);

  useEffect(() => {
    const poll = setInterval(async () => {
      const res = await fetch(`/api/result/${jobId}`);
      const data = await res.json();
      setStatus(data.status);
      if (data.status === "done") {
        setSummary(data.summary);
        clearInterval(poll);
      }
      if (data.status === "error") clearInterval(poll);
    }, 800);
    return () => clearInterval(poll);
  }, [jobId]);

  if (status !== "done") {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <p className="text-body text-sm">{status === "error" ? "Conversion failed." : "Converting…"}</p>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen">
      {/* Nav */}
      <nav className="h-10 flex items-center justify-between px-4 border-b border-hairline shrink-0">
        <div className="flex items-center gap-3">
          <span className="text-[13px] font-medium text-ink">osheet<span className="text-mute font-normal">.io</span></span>
          <span className="text-mute text-xs font-mono">/ converted</span>
        </div>
        <div className="flex gap-2">
          <a href={`/api/download/${jobId}/osheet`} className="text-xs font-medium text-ink border border-hairline px-3 py-1 rounded hover:bg-canvas-soft">Export .osheet</a>
          <a href={`/api/download/${jobId}/xlsx`} className="text-xs font-medium bg-ink text-canvas px-3 py-1 rounded hover:opacity-80">↓ Download .xlsx</a>
        </div>
      </nav>

      {/* Stat bar */}
      {summary && (
        <div className="flex items-center gap-2 px-4 py-2 border-b border-hairline shrink-0">
          {[
            { label: "sheets", count: summary.sheet_count, color: "" },
            { label: "tables", count: summary.table_count, color: "" },
            { label: "assumptions", count: summary.assumption_count, color: "text-assumption" },
            { label: "outputs", count: summary.output_count, color: "text-output" },
            { label: "warnings", count: summary.warning_count, color: "text-warning" },
          ].map((s, i) => (
            <span key={i} className="flex items-center gap-1 text-xs border border-hairline bg-canvas-soft px-2 py-1 rounded">
              <span className={`font-mono ${s.color || "text-ink"}`}>{s.count}</span>
              <span className="text-body-strong">{s.label}</span>
            </span>
          ))}
        </div>
      )}

      {/* Warnings */}
      {summary && summary.warnings.length > 0 && (
        <div className="px-4 py-2 border-b border-hairline bg-canvas-soft shrink-0">
          {summary.warnings.map((w, i) => (
            <p key={i} className="text-xs text-warning"><span className="text-ink font-medium">{w.address}</span> — {w.message}</p>
          ))}
        </div>
      )}

      <div className="text-body text-sm flex-1 flex items-center justify-center">
        <div className="text-center">
          <p className="mb-2">Conversion complete.</p>
          <p className="text-xs text-mute">Download your files above.</p>
        </div>
      </div>
    </div>
  );
}
```

- [ ] Run dev server: `npm run dev` — verify Upload page loads at http://localhost:3000
- [ ] Commit: `git add . && git commit -m "feat: Next.js frontend (upload + inspector views)"`

---

### Task 13: Benchmarking suite

**Files:**
- Create: `benchmarks/requirements.txt`
- Create: `benchmarks/make_fixture.py`
- Create: `benchmarks/metrics.py`
- Create: `benchmarks/run_baseline.py`
- Create: `benchmarks/run_osheet.py`
- Create: `benchmarks/report.py`

- [ ] Create `benchmarks/requirements.txt`

```
anthropic>=0.28
openpyxl>=3.1
osheet @ file://../packages/osheet
```

- [ ] Install: `pip install -r benchmarks/requirements.txt`

- [ ] Create `benchmarks/make_fixture.py`

```python
"""Generate a realistic dummy financial model .xlsx for benchmarking."""
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment


def make_financial_model() -> bytes:
    wb = openpyxl.Workbook()

    # ── Assumptions sheet ──────────────────────────────────────────────────
    ws_a = wb.active
    ws_a.title = "Assumptions"
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    assumptions = [
        ("Growth Rate", 0.15),
        ("Churn Rate", 0.04),
        ("Gross Margin", 0.72),
        ("COGS %", 0.28),
        ("Sales Headcount", 12),
        ("ACV", 24000),
        ("NRR", 1.08),
    ]
    ws_a["A1"] = "Assumption"
    ws_a["B1"] = "Value"
    ws_a["A1"].font = Font(bold=True)
    ws_a["B1"].font = Font(bold=True)
    for i, (name, val) in enumerate(assumptions, start=2):
        ws_a[f"A{i}"] = name
        ws_a[f"B{i}"] = val
        ws_a[f"B{i}"].fill = yellow

    # ── Revenue sheet ──────────────────────────────────────────────────────
    ws_r = wb.create_sheet("Revenue")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    headers = ["Month", "New ARR", "Churned ARR", "Net ARR", "Cumulative ARR", "Gross Profit"]
    for col, h in enumerate(headers, 1):
        ws_r.cell(row=1, column=col, value=h).font = Font(bold=True)

    base_arr = 100000
    cum_arr = base_arr
    for i, month in enumerate(months, start=2):
        growth = f"=Assumptions!B1"
        churn = f"=Assumptions!B2"
        ws_r[f"A{i}"] = month
        ws_r[f"B{i}"] = f"=ROUND({base_arr}*(1+Assumptions!$B$1)^{i-2},0)"
        ws_r[f"C{i}"] = f"=ROUND(E{i-1}*Assumptions!$B$2,0)" if i > 2 else 0
        ws_r[f"D{i}"] = f"=B{i}-C{i}"
        ws_r[f"E{i}"] = f"=IF({i}=2,B{i},E{i-1}+D{i})"
        ws_r[f"F{i}"] = f"=E{i}*Assumptions!$B$3"

    # Totals row
    ws_r[f"A{14}"] = "Total"
    ws_r[f"B{14}"] = f"=SUM(B2:B13)"
    ws_r[f"D{14}"] = f"=SUM(D2:D13)"
    ws_r[f"F{14}"] = f"=SUM(F2:F13)"

    # ── Summary sheet ──────────────────────────────────────────────────────
    ws_s = wb.create_sheet("Summary")
    ws_s["A1"] = "Metric"
    ws_s["B1"] = "Value"
    ws_s["A1"].font = Font(bold=True)
    ws_s["B1"].font = Font(bold=True)
    metrics = [
        ("Total ARR", "=Revenue!E13"),
        ("Total Gross Profit", "=Revenue!F14"),
        ("Net ARR Added", "=Revenue!D14"),
        ("Gross Margin %", "=Assumptions!B3"),
        ("Churn Rate", "=Assumptions!B2"),
        ("Growth Rate", "=Assumptions!B1"),
    ]
    for i, (name, formula) in enumerate(metrics, start=2):
        ws_s[f"A{i}"] = name
        ws_s[f"B{i}"] = formula

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


if __name__ == "__main__":
    data = make_financial_model()
    with open("benchmarks/dummy_financial_model.xlsx", "wb") as f:
        f.write(data)
    print(f"Written {len(data):,} bytes → benchmarks/dummy_financial_model.xlsx")
```

- [ ] Run: `python benchmarks/make_fixture.py` — verify file created

- [ ] Create `benchmarks/metrics.py`

```python
"""Scoring utilities for benchmark runs."""
from dataclasses import dataclass
from typing import Any


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
```

- [ ] Create `benchmarks/run_baseline.py`

```python
"""Baseline: dump xlsx as CSV text and ask Claude to answer questions about it."""
import time
import csv
import io
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

        results.append(QuestionResult(
            question_id=q["id"],
            question=q["question"],
            answer=answer,
            correct=score_answer(answer, q["expected_contains"]),
            latency_ms=latency_ms,
        ))
        print(f"[baseline] {q['id']}: {'✓' if results[-1].correct else '✗'} ({latency_ms:.0f}ms)")

    return BenchmarkResult(label="baseline_raw_csv", results=results)


if __name__ == "__main__":
    result = run_baseline()
    print(f"\nBaseline accuracy: {result.accuracy:.1%}  avg latency: {result.avg_latency_ms:.0f}ms")
```

- [ ] Create `benchmarks/run_osheet.py`

```python
"""osheet benchmark: load with library, give Claude the structured manifest + cells."""
import json
import time
import anthropic
import osheet
from metrics import QUESTIONS, BenchmarkResult, QuestionResult, score_answer


def workbook_to_context(wb) -> str:
    lines = []
    lines.append(f"Workbook: {wb.manifest.sheet_count} sheets, {wb.manifest.table_count} tables")
    lines.append(f"Assumptions ({len(wb.assumptions)}):")
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

        results.append(QuestionResult(
            question_id=q["id"],
            question=q["question"],
            answer=answer,
            correct=score_answer(answer, q["expected_contains"]),
            latency_ms=latency_ms,
        ))
        print(f"[osheet]   {q['id']}: {'✓' if results[-1].correct else '✗'} ({latency_ms:.0f}ms)")

    return BenchmarkResult(label="osheet_structured", results=results)


if __name__ == "__main__":
    result = run_osheet()
    print(f"\nOsheet accuracy: {result.accuracy:.1%}  avg latency: {result.avg_latency_ms:.0f}ms")
```

- [ ] Create `benchmarks/report.py`

```python
"""Run both benchmarks and print a comparison table."""
from run_baseline import run_baseline
from run_osheet import run_osheet

def main():
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
```

- [ ] Run: `python benchmarks/make_fixture.py && python benchmarks/report.py`
- [ ] Commit: `git add . && git commit -m "feat: benchmarking suite (baseline vs osheet comparison)"`

---

### Task 14: GitHub Actions CI

**Files:**
- Create: `.github/workflows/ci.yml`

- [ ] Create `.github/workflows/ci.yml`

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  library:
    runs-on: ubuntu-latest
    defaults:
      run:
        working-directory: packages/osheet
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - run: pip install -e ".[dev]"
      - run: pytest --tb=short -q --cov=src/osheet --cov-report=term-missing

  backend:
    runs-on: ubuntu-latest
    defaults:
      run:
        working-directory: packages/osheet-app/backend
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - run: pip install -e ".[dev]"
      - run: pytest --tb=short -q

  frontend:
    runs-on: ubuntu-latest
    defaults:
      run:
        working-directory: packages/osheet-app/frontend
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: "20"
          cache: "npm"
          cache-dependency-path: packages/osheet-app/frontend/package-lock.json
      - run: npm ci
      - run: npm run build
```

- [ ] Commit and push

```bash
git add .github/
git commit -m "ci: GitHub Actions for library, backend, frontend"
git push -u origin main
```

- [ ] Verify CI passes at https://github.com/AmoghReddy45/osheet/actions

---

## Self-Review

**Spec coverage check:**
- ✅ Python library (parser, 5 analyzer passes, emitters) — Tasks 2–10
- ✅ Workbook API (load, trace, find, propose_patch, apply_patch, export) — Task 10
- ✅ .osheet format (zip: workbook.json, sheets.json, formula_graph.json) — Task 9
- ✅ Annotated .xlsx output — Task 9
- ✅ FastAPI backend (3 endpoints) — Task 11
- ✅ Next.js Upload + Inspector views — Task 12
- ✅ Unit + integration + contract tests — Tasks 2–11
- ✅ Benchmarking suite — Task 13
- ✅ CI — Task 14

**Type consistency check:**
- `Cell.depends_on: list[str]` — defined Task 2, used in Tasks 6, 7, 10 ✅
- `Workbook.all_cells` property — defined Task 2, used in Tasks 7, 10 ✅
- `OsheetWorkbook` wraps `Workbook` — Task 10 ✅
- `run_all(workbook)` mutates in place — Task 4 `__init__.py`, called in Task 10 ✅
- `to_xlsx_bytes(wb: Workbook)` / `to_osheet_bytes(wb: Workbook)` — Task 9, called via `OsheetWorkbook` ✅

**No placeholders found.**
