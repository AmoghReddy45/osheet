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
