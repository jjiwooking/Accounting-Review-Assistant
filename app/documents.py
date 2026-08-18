from __future__ import annotations

import re
import shutil
import uuid
from pathlib import Path

from docx import Document
from pypdf import PdfReader

from .db import BASE_DIR, db
from .services import now_iso, validate_all

UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
ALLOWED = {".pdf", ".docx", ".txt", ".jpg", ".jpeg", ".png"}
MAX_BYTES = 15 * 1024 * 1024

CATEGORY_KEYWORDS = {
    "세금계산서": ["세금계산서", "공급가액", "부가가치세", "사업자등록번호"],
    "영수증": ["영수증", "승인번호", "카드", "가맹점", "receipt"],
    "거래명세서": ["거래명세", "거래명세서", "품목", "수량", "단가"],
    "계약서": ["계약서", "계약금액", "계약기간", "갑", "을"],
    "견적서": ["견적서", "견적금액", "quotation", "estimate"],
    "지출결의": ["지출결의", "품의", "승인", "결재"],
}


def _safe_name(filename: str) -> str:
    base = Path(filename).name
    stem = re.sub(r"[^0-9A-Za-z가-힣._-]+", "_", Path(base).stem)[:80] or "document"
    return f"{uuid.uuid4().hex[:12]}_{stem}{Path(base).suffix.lower()}"


def extract_text(path: Path) -> str:
    suffix = path.suffix.lower()
    try:
        if suffix == ".pdf":
            reader = PdfReader(str(path))
            parts = []
            for page in reader.pages[:25]:
                text = page.extract_text() or ""
                if text:
                    parts.append(text)
                if sum(len(p) for p in parts) > 80_000:
                    break
            return "\n".join(parts)[:100_000]
        if suffix == ".docx":
            doc = Document(str(path))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())[:100_000]
        if suffix == ".txt":
            raw = path.read_bytes()
            for enc in ("utf-8-sig", "cp949", "euc-kr"):
                try:
                    return raw.decode(enc)[:100_000]
                except UnicodeDecodeError:
                    continue
    except Exception:
        return ""
    return ""


def classify(filename: str, text: str) -> str:
    haystack = f"{filename}\n{text}".lower()
    scores = {}
    for category, words in CATEGORY_KEYWORDS.items():
        scores[category] = sum(haystack.count(w.lower()) for w in words)
    best = max(scores, key=scores.get)
    return best if scores[best] > 0 else "기타"


def summarize(text: str, category: str) -> str:
    if not text.strip():
        return f"{category} 문서로 분류되었습니다. 이미지 문서는 현재 본문 OCR 자동확정을 수행하지 않으므로 원본 확인이 필요합니다."
    cleaned = re.sub(r"\s+", " ", text).strip()
    # Deterministic extractive summary: first meaningful sentences/phrases only, no invented facts.
    sentences = re.split(r"(?<=[.!?다요])\s+|\n+", cleaned)
    selected = []
    for sentence in sentences:
        s = sentence.strip()
        if len(s) >= 12:
            selected.append(s)
        if len(" ".join(selected)) >= 350 or len(selected) >= 3:
            break
    body = " ".join(selected)[:450] if selected else cleaned[:450]
    return f"자동 분류: {category}. 원문 발췌 요약: {body}"


def save_document(filename: str, content: bytes, reference_no: str = "", category_override: str = "") -> int:
    suffix = Path(filename).suffix.lower()
    if suffix not in ALLOWED:
        raise ValueError("지원 문서: PDF, DOCX, TXT, JPG, JPEG, PNG")
    if len(content) > MAX_BYTES:
        raise ValueError("문서 크기는 15MB 이하만 업로드할 수 있습니다.")
    stored = _safe_name(filename)
    path = UPLOAD_DIR / stored
    path.write_bytes(content)
    text = extract_text(path)
    category = category_override.strip() or classify(filename, text)
    summary = summarize(text, category)
    with db() as conn:
        cur = conn.execute(
            "INSERT INTO documents(filename,stored_name,category,reference_no,summary,text_preview,uploaded_at) VALUES(?,?,?,?,?,?,?)",
            (Path(filename).name, stored, category, reference_no.strip() or None, summary, text[:5000] or None, now_iso()),
        )
        doc_id = cur.lastrowid
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,?)", ("문서 업로드", "document", str(doc_id), f"{filename} / {category}", now_iso()))
    # Evidence linkage can change review findings, so re-run rules after a document arrives.
    validate_all()
    return int(doc_id)


def list_documents() -> list[dict]:
    with db() as conn:
        return [dict(r) for r in conn.execute("SELECT * FROM documents ORDER BY id DESC")]


def get_document(doc_id: int) -> dict | None:
    with db() as conn:
        row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
        return dict(row) if row else None


def delete_document(doc_id: int) -> None:
    with db() as conn:
        row = conn.execute("SELECT stored_name, filename FROM documents WHERE id=?", (doc_id,)).fetchone()
        if not row:
            return
        conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
        conn.execute("INSERT INTO audit_log(action,target_type,target_id,detail,created_at) VALUES(?,?,?,?,?)", ("문서 삭제", "document", str(doc_id), row["filename"], now_iso()))
    path = UPLOAD_DIR / row["stored_name"]
    if path.exists():
        path.unlink()
    validate_all()
