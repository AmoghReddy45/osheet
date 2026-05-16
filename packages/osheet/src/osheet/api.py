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
        downstream = [c.stable_id for c in self._wb.all_cells if stable_id in c.depends_on]
        return TraceResult(cell_id=stable_id, upstream=cell.depends_on, downstream=downstream)

    def propose_patch(self, stable_id: str, new_value: object) -> PatchProposal:
        cell = self._wb.get_cell(stable_id)
        if cell is None:
            return PatchProposal(cell_id=stable_id, old_value=None, new_value=new_value)
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
