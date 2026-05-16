from __future__ import annotations
import time
import uuid
from dataclasses import dataclass, field

_TTL_SECONDS = 3600


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
