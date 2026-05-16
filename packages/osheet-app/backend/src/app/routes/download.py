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
        return Response(
            content=job.xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=output.ai.xlsx"}
        )
    if file_type == "osheet":
        return Response(
            content=job.osheet_bytes,
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=output.osheet"}
        )
    raise HTTPException(status_code=400, detail="file_type must be xlsx or osheet")
