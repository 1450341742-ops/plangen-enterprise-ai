from __future__ import annotations

from pathlib import Path
from docx import Document


def docx_to_markdown_text(docx_path: str | Path) -> str:
    doc = Document(str(docx_path))
    parts: list[str] = []
    for p in doc.paragraphs:
        if p.text.strip():
            parts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip().replace("\n", " ") for c in row.cells]
            if any(cells):
                parts.append("| " + " | ".join(cells) + " |")
    return "\n".join(parts)


def pdf_to_text(pdf_path: str | Path) -> str:
    """Extract text from normal text PDFs. Scanned PDFs may need OCR and should be converted or uploaded as DOCX."""
    pdf_path = Path(pdf_path)
    try:
        import fitz  # PyMuPDF
    except Exception as e:
        raise RuntimeError("缺少PDF解析依赖 PyMuPDF，请在 requirements.txt 中加入 PyMuPDF") from e

    parts: list[str] = []
    doc = fitz.open(str(pdf_path))
    for page_no, page in enumerate(doc, 1):
        text = page.get_text("text") or ""
        if text.strip():
            parts.append(f"\n\n--- PDF第{page_no}页 ---\n{text.strip()}")
    return "\n".join(parts).strip()


def uploaded_file_to_text(file_path: str | Path) -> str:
    file_path = Path(file_path)
    suffix = file_path.suffix.lower()
    if suffix == ".docx":
        return docx_to_markdown_text(file_path)
    if suffix == ".pdf":
        return pdf_to_text(file_path)
    if suffix in [".txt", ".md"]:
        return file_path.read_text(encoding="utf-8", errors="ignore")
    raise ValueError(f"暂不支持的文件格式：{suffix}")
