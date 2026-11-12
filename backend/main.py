import io
import os
import time
import uuid
import zipfile
import logging
import asyncio
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Request
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

from processor import process_pdf_bytes, build_zip, _check_ocr

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

JOBS: dict[str, dict] = {}

# Limits
MAX_UPLOAD_BYTES = 100 * 1024 * 1024       # 100 MB por arquivo
MAX_FILES_PER_REQUEST = 20                  # máximo de arquivos por envio
MAX_ZIP_MEMBERS = 50                        # máximo de PDFs dentro de um ZIP
MAX_ZIP_EXTRACTED_BYTES = 300 * 1024 * 1024 # 300 MB total extraído do ZIP
JOB_TTL_SECONDS = 3600                      # jobs expiram após 1 hora

# Magic bytes válidos
_PDF_MAGIC = b"%PDF"
_ZIP_MAGIC = (b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08")

ALLOWED_ORIGINS = os.getenv("ALLOWED_ORIGINS", "http://localhost:5173").split(",")

app = FastAPI(title="DS OCR Unifier", version="1.0.0", docs_url=None, redoc_url=None)

app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


def _validate_magic(raw: bytes, filename: str) -> None:
    """Rejeita arquivos cujo conteúdo não corresponde à extensão declarada."""
    if filename.lower().endswith(".pdf"):
        if not raw[:4] == _PDF_MAGIC:
            raise HTTPException(400, f"'{filename}' não é um PDF válido.")
    elif filename.lower().endswith(".zip"):
        if raw[:4] not in _ZIP_MAGIC:
            raise HTTPException(400, f"'{filename}' não é um ZIP válido.")


def _cleanup_old_jobs() -> None:
    """Remove jobs expirados da memória, incluindo os bytes dos PDFs gerados."""
    now = time.time()
    stale = [jid for jid, j in JOBS.items() if now - j.get("created_at", now) > JOB_TTL_SECONDS]
    for jid in stale:
        JOBS.pop(jid, None)
    if stale:
        logger.info("Limpeza: %d job(s) expirado(s) removidos", len(stale))


def _collect_pdfs_from_zip(zip_bytes: bytes, zip_filename: str) -> list[tuple[str, bytes]]:
    """Extrai PDFs de um ZIP com proteção contra ZIP bomb."""
    collected = []
    total_bytes = 0
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            members = [i for i in zf.infolist() if i.filename.lower().endswith(".pdf") and not i.is_dir()]
            if len(members) > MAX_ZIP_MEMBERS:
                raise HTTPException(400, f"ZIP contém mais de {MAX_ZIP_MEMBERS} PDFs.")
            for info in members:
                total_bytes += info.file_size
                if total_bytes > MAX_ZIP_EXTRACTED_BYTES:
                    raise HTTPException(400, "ZIP excede o limite de 300 MB extraído.")
                data = zf.read(info.filename)
                name = Path(info.filename).name  # remove qualquer ../ no path
                if data[:4] == _PDF_MAGIC:        # valida magic byte de cada PDF interno
                    collected.append((name, data))
                else:
                    logger.warning("ZIP '%s': membro '%s' ignorado (não é PDF válido)", zip_filename, name)
    except zipfile.BadZipFile:
        raise HTTPException(400, f"'{zip_filename}' está corrompido ou não é um ZIP válido.")
    return collected


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
        job["error"] = "Erro interno no processamento. Verifique se o PDF é válido e tente novamente."


@app.get("/api/health")
async def health():
    return {"status": "ok", "ocr_available": _check_ocr()}


@app.post("/api/process")
async def process_files(background_tasks: BackgroundTasks, files: list[UploadFile] = File(...)):
    _cleanup_old_jobs()

    if len(files) > MAX_FILES_PER_REQUEST:
        raise HTTPException(400, f"Máximo de {MAX_FILES_PER_REQUEST} arquivos por envio.")

    files_data: list[tuple[str, bytes]] = []

    for upload in files:
        raw = await upload.read()
        if len(raw) > MAX_UPLOAD_BYTES:
            raise HTTPException(413, f"'{upload.filename}' excede o limite de 100 MB.")

        fname = (upload.filename or "upload.pdf").strip()
        _validate_magic(raw, fname)

        if fname.lower().endswith(".zip"):
            pdfs = _collect_pdfs_from_zip(raw, fname)
            if not pdfs:
                raise HTTPException(400, f"ZIP '{fname}' não contém PDFs válidos.")
            files_data.extend(pdfs)
        elif fname.lower().endswith(".pdf"):
            files_data.append((fname, raw))
        else:
            raise HTTPException(400, f"Tipo de arquivo não suportado: '{fname}'. Use PDF ou ZIP.")

    if not files_data:
        raise HTTPException(400, "Nenhum PDF encontrado no envio.")

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status": "pending", "progress": 0, "results": [],
        "zip_bytes": None, "error": "", "created_at": time.time(),
    }
    background_tasks.add_task(_process_job, job_id, files_data)
    return {"job_id": job_id, "file_count": len(files_data)}


@app.get("/api/status/{job_id}")
async def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado.")
    return {"job_id": job_id, "status": job["status"], "progress": job["progress"],
            "results": job.get("results", []), "error": job.get("error", "")}


@app.get("/api/download/{job_id}")
async def download_zip(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job não encontrado.")
    if job["status"] != "done":
        raise HTTPException(400, "Processamento não concluído.")
    zip_data = job.get("zip_bytes")
    if not zip_data:
        raise HTTPException(500, "Arquivo ZIP não disponível.")
    return StreamingResponse(
        io.BytesIO(zip_data),
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="DS_Unificadas.zip"'},
    )
