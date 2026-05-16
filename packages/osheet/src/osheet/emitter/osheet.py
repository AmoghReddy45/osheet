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
