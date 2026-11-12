import io
import os
import uuid
import zipfile
import logging
import asyncio
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

from processor import process_pdf_bytes, build_zip, _check_ocr

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

JOBS: dict[str, dict] = {}
MAX_UPLOAD_BYTES = 200 * 1024 * 1024

app = FastAPI(title="DS OCR Unifier", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


async def _process_job(job_id: str, files_data: list[tuple[str, bytes]]) -> None:
    job = JOBS[job_id]
    job["status"] = "processing"
    job["progress"] = 0
    all_results: list[dict] = []
    total = len(files_data)

    try:
        for idx, (filename, pdf_bytes) in enumerate(files_data):
            loop = asyncio.get_event_loop()
            results = await loop.run_in_executor(None, process_pdf_bytes, pdf_bytes, filename)
            all_results.extend(results)
            job["progress"] = int((idx + 1) / total * 90)

        job["results"] = [{k: v for k, v in r.items() if k != "pdf_bytes"} for r in all_results]

        zip_bytes = await asyncio.get_event_loop().run_in_executor(None, build_zip, all_results)
        job["zip_bytes"] = zip_bytes
        job["progress"] = 100
        job["status"] = "done"

    except Exception as exc:
        logger.exception("[%s] Processing failed: %s", job_id, exc)
        job["status"] = "error"
        job["error"] = str(exc)


def _collect_pdfs_from_zip(zip_bytes: bytes) -> list[tuple[str, bytes]]:
    collected = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            if info.filename.lower().endswith(".pdf") and not info.is_dir():
                collected.append((Path(info.filename).name, zf.read(info.filename)))
    return collected


@app.get("/api/health")
async def health():
    return {"status": "ok", "ocr_available": _check_ocr()}


@app.post("/api/process")
async def process_files(background_tasks: BackgroundTasks, files: list[UploadFile] = File(...)):
    files_data: list[tuple[str, bytes]] = []

    for upload in files:
        raw = await upload.read()
        if len(raw) > MAX_UPLOAD_BYTES:
            raise HTTPException(413, f"File '{upload.filename}' exceeds 200 MB limit")

        fname = upload.filename or "upload.pdf"
        if fname.lower().endswith(".zip"):
            pdfs = _collect_pdfs_from_zip(raw)
            if not pdfs:
                raise HTTPException(400, f"ZIP '{fname}' contains no PDF files")
            files_data.extend(pdfs)
        elif fname.lower().endswith(".pdf"):
            files_data.append((fname, raw))
        else:
            raise HTTPException(400, f"Unsupported file type: '{fname}'")

    if not files_data:
        raise HTTPException(400, "No PDF files found in the upload")

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status": "pending", "progress": 0, "results": [], "zip_bytes": None, "error": ""}
    background_tasks.add_task(_process_job, job_id, files_data)
    return {"job_id": job_id, "file_count": len(files_data)}


@app.get("/api/status/{job_id}")
async def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    return {"job_id": job_id, "status": job["status"], "progress": job["progress"],
            "results": job.get("results", []), "error": job.get("error", "")}


@app.get("/api/download/{job_id}")
async def download_zip(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job not found")
    if job["status"] != "done":
        raise HTTPException(400, "Job not finished yet")
    return StreamingResponse(
        io.BytesIO(job["zip_bytes"]),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="DS_Unificadas.zip"'},
    )
