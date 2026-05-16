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
