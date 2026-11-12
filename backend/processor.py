"""
PDF processing module: extracts client info, groups pages, merges and renames PDFs.
Uses direct text extraction first, falls back to OCR via tesseract/pdf2image.
"""

import re
import io
import os
import zipfile
import logging
from typing import Optional

from pypdf import PdfReader, PdfWriter
from PIL import Image

logger = logging.getLogger(__name__)

MONTHS_PT = {
    "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL",
    "05": "MAIO",    "06": "JUNHO",     "07": "JULHO",  "08": "AGOSTO",
    "09": "SETEMBRO","10": "OUTUBRO",   "11": "NOVEMBRO","12": "DEZEMBRO",
}

_NAME_SUFFIXES = [
    " DO BRASIL", " LTDA", " S/A", " S.A.", " S.A",
    " COMERCIAL E PARTICIPACOES", " COMERCIAL", " INDUSTRIA",
    " IND. E COMERCIO", " E COMERCIO", " EIRELI", " ME", " EPP",
]

_OCR_AVAILABLE: Optional[bool] = None


def _check_ocr() -> bool:
    global _OCR_AVAILABLE
    if _OCR_AVAILABLE is not None:
        return _OCR_AVAILABLE
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        _OCR_AVAILABLE = True
    except Exception:
        _OCR_AVAILABLE = False
    return _OCR_AVAILABLE


def clean_client_name(name: str) -> str:
    name = name.split("\n")[0]
    name = re.split(
        r"\s+(?:Endere[çc]o|ROD\.|AV\.|RUA|CIDADE|CNPJ|CPF)\b",
        name, maxsplit=1, flags=re.IGNORECASE
    )[0]
    name = re.sub(r"[\d\-/,.]+$", "", name).strip()
    for suffix in _NAME_SUFFIXES:
        idx = name.upper().find(suffix)
        if idx != -1:
            name = name[:idx]
    name = name.strip(" ,.-/")
    if len(name) > 80:
        name = name[:80].rstrip()
    return name.upper()


def get_month_name(month_num: str) -> str:
    return MONTHS_PT.get(month_num.zfill(2), "MES_DESCONHECIDO")


def extract_info_from_text(text: str) -> dict:
    info = {}
    client_pat = re.compile(
        r"Cliente[.:]+\s*(\d{5})-\d{3}-\d{2}\s+([^\n]+)",
        re.IGNORECASE,
    )
    m = client_pat.search(text)
    if m:
        raw_code = m.group(1)
        info["code"] = str(int(raw_code))
        info["full_name"] = m.group(2).strip()
        info["client_name"] = clean_client_name(info["full_name"])

    date_pat = re.compile(r"Mes/Ano[.:]+\s*(\d{2})/(\d{4})", re.IGNORECASE)
    d = date_pat.search(text)
    if d:
        info["month"] = d.group(1)
        info["year"] = d.group(2)
        info["month_name"] = get_month_name(info["month"])

    return info


def _ocr_page(page) -> str:
    try:
        from pdf2image import convert_from_bytes
        import pytesseract

        writer = PdfWriter()
        writer.add_page(page)
        buf = io.BytesIO()
        writer.write(buf)
        buf.seek(0)

        images = convert_from_bytes(buf.read(), dpi=200)
        texts = []
        for img in images:
            txt = pytesseract.image_to_string(img, lang="por+eng")
            texts.append(txt)
        return "\n".join(texts)
    except Exception as exc:
        logger.warning("OCR failed for page: %s", exc)
        return ""


def extract_text_from_page(page) -> str:
    text = page.extract_text() or ""
    if len(text.strip()) < 50 and _check_ocr():
        text = _ocr_page(page)
    return text


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "", name).strip()


def build_output_name(client_name: str, code: str, month_name: str, year: str) -> str:
    base = f"{client_name} {code} - COMPLEMENTAR - {month_name} {year}"
    return safe_filename(base) + ".pdf"


def process_pdf_bytes(pdf_bytes: bytes, source_filename: str) -> list[dict]:
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception as exc:
        logger.error("Cannot open PDF '%s': %s", source_filename, exc)
        return []

    groups: dict[str, dict] = {}

    for page_num, page in enumerate(reader.pages):
        text = extract_text_from_page(page)
        info = extract_info_from_text(text)

        if not info.get("code"):
            logger.warning("'%s' page %d: no client code found – skipping",
                           source_filename, page_num + 1)
            continue

        code = info["code"]
        if code not in groups:
            groups[code] = {
                "client_name": info.get("client_name", "DESCONHECIDO"),
                "month_name": info.get("month_name", "MES_DESCONHECIDO"),
                "year": info.get("year", "0000"),
                "pages": [],
            }
        groups[code]["pages"].append(page)

        if info.get("client_name") and groups[code]["client_name"] == "DESCONHECIDO":
            groups[code]["client_name"] = info["client_name"]
        if info.get("month_name") and groups[code]["month_name"] == "MES_DESCONHECIDO":
            groups[code]["month_name"] = info["month_name"]
            groups[code]["year"] = info.get("year", "0000")

    results = []
    for code, data in groups.items():
        writer = PdfWriter()
        for p in data["pages"]:
            writer.add_page(p)

        buf = io.BytesIO()
        writer.write(buf)
        pdf_out = buf.getvalue()

        filename = build_output_name(
            data["client_name"], code, data["month_name"], data["year"]
        )
        results.append({
            "filename": filename,
            "pdf_bytes": pdf_out,
            "client_name": data["client_name"],
            "code": code,
            "month_name": data["month_name"],
            "year": data["year"],
            "page_count": len(data["pages"]),
        })

    return results


def build_zip(processed: list[dict]) -> bytes:
    buf = io.BytesIO()
    seen: dict[str, int] = {}

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in processed:
            name = item["filename"]
            if name in seen:
                seen[name] += 1
                base, ext = os.path.splitext(name)
                name = f"{base} ({seen[name]}){ext}"
            else:
                seen[name] = 0
            zf.writestr(name, item["pdf_bytes"])

    return buf.getvalue()
