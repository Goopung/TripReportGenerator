import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from docx import Document
from pypdf import PdfReader


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp"}
PDF_EXTS = {".pdf"}
DOCX_EXTS = {".docx"}
TEXT_EXTS = {".txt", ".md"}


def safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:180]


def slugify(text: str) -> str:
    text = text or "Trip_Report"
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text)
    text = re.sub(r"\s+", "_", text.strip())
    return text.strip("_")[:120] or "Trip_Report"


def iso_date(value: date | datetime | str) -> str:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def korean_date(value: date | datetime | str) -> str:
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value.strftime("%Y년 %m월 %d일")
    try:
        parsed = datetime.fromisoformat(str(value)).date()
        return parsed.strftime("%Y년 %m월 %d일")
    except Exception:
        return str(value)


def date_range(start: date, end: date) -> list[date]:
    if end < start:
        return []
    out: list[date] = []
    current = start
    while current <= end:
        out.append(current)
        current += timedelta(days=1)
    return out


def save_uploaded_file(uploaded_file: Any, target_dir: str | Path, prefix: str = "") -> str:
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)

    base_name = safe_filename(uploaded_file.name)
    if prefix:
        base_name = f"{safe_filename(prefix)}_{base_name}"

    target = target_dir / base_name
    count = 1
    while target.exists():
        target = target_dir / f"{target.stem}_{count}{target.suffix}"
        count += 1

    with open(target, "wb") as f:
        f.write(uploaded_file.getbuffer())

    return str(target)


def save_uploaded_files(uploaded_files: Any, target_dir: str | Path, prefix: str = "") -> list[str]:
    if uploaded_files is None:
        return []
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]

    paths = []
    for uploaded_file in uploaded_files:
        if uploaded_file is not None:
            paths.append(save_uploaded_file(uploaded_file, target_dir, prefix))
    return paths


def extract_text_from_file(path: str | Path, max_chars: int = 16000) -> str:
    path = Path(path)
    suffix = path.suffix.lower()

    try:
        if suffix in TEXT_EXTS:
            return path.read_text(encoding="utf-8", errors="ignore")[:max_chars]

        if suffix in PDF_EXTS:
            reader = PdfReader(str(path))
            texts = []
            for page in reader.pages[:30]:
                texts.append(page.extract_text() or "")
            return "\n".join(texts)[:max_chars]

        if suffix in DOCX_EXTS:
            doc = Document(str(path))
            texts = [p.text for p in doc.paragraphs if p.text.strip()]
            return "\n".join(texts)[:max_chars]
    except Exception as exc:
        return f"[텍스트 추출 실패: {exc}]"

    return ""


def collect_paths(value: Any) -> list[str]:
    paths: list[str] = []

    def walk(x: Any) -> None:
        if isinstance(x, str) and x:
            paths.append(x)
        elif isinstance(x, list):
            for item in x:
                walk(item)
        elif isinstance(x, dict):
            for item in x.values():
                walk(item)

    walk(value)
    return [p for p in paths if Path(p).exists()]
