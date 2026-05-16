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
    if len(data) > 20 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File too large (max 20 MB)")
    job = create_job()
    background_tasks.add_task(_run_conversion, job, data)
    return {"job_id": job.job_id, "status": "processing"}
