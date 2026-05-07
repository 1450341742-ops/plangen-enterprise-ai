from __future__ import annotations

import math
import re
from pathlib import Path
from typing import List, Tuple

BASE_DIR = Path(__file__).resolve().parents[2]
KB_DIR = BASE_DIR / "knowledge_base"


def _read_kb_files() -> list[tuple[str, str]]:
    if not KB_DIR.exists():
        return []
    items: list[tuple[str, str]] = []
    for ext in ("*.md", "*.txt"):
        for path in sorted(KB_DIR.glob(ext)):
            try:
                text = path.read_text(encoding="utf-8", errors="ignore")
                if text.strip():
                    items.append((path.name, text))
            except Exception:
                continue
    return items


def _split_chunks(text: str, chunk_size: int = 900, overlap: int = 120) -> list[str]:
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    if not text:
        return []
    parts = re.split(r"(?=\n#{1,4}\s+|\n\d+(?:\.\d+)*\s+|\n[一二三四五六七八九十]+、)", text)
    chunks: list[str] = []
    for part in parts:
        p = part.strip()
        if not p:
            continue
        if len(p) <= chunk_size:
            chunks.append(p)
        else:
            start = 0
            while start < len(p):
                chunks.append(p[start:start + chunk_size])
                start += max(1, chunk_size - overlap)
    return chunks


def _tokens(text: str) -> list[str]:
    text = text.lower()
    words = re.findall(r"[a-zA-Z0-9_\-]+|[\u4e00-\u9fff]{1,2}", text)
    return [w for w in words if w.strip()]


def _score(query_tokens: list[str], chunk: str) -> float:
    if not query_tokens:
        return 0.0
    c_tokens = _tokens(chunk)
    if not c_tokens:
        return 0.0
    freq = {}
    for t in c_tokens:
        freq[t] = freq.get(t, 0) + 1
    score = 0.0
    for qt in query_tokens:
        score += freq.get(qt, 0)
    # 关键临床质控词加权
    boost_terms = ["入排", "入选", "排除", "rbqm", "风险", "法规", "gcp", "ich", "edc", "icf", "sae", "药品", "样本", "实验室", "时间窗", "源数据", "一致性"]
    for term in boost_terms:
        if term.lower() in chunk.lower():
            score += 0.8
    return score / math.sqrt(len(c_tokens))


def retrieve_knowledge(query: str, top_k: int = 6) -> str:
    files = _read_kb_files()
    if not files:
        return ""
    q_tokens = _tokens(query)
    scored: list[tuple[float, str, str]] = []
    for filename, text in files:
        for chunk in _split_chunks(text):
            s = _score(q_tokens, chunk)
            if s > 0:
                scored.append((s, filename, chunk))
    scored.sort(key=lambda x: x[0], reverse=True)
    selected = scored[:top_k]
    if not selected:
        return ""
    parts = []
    for i, (score, filename, chunk) in enumerate(selected, 1):
        parts.append(f"【知识片段{i}｜{filename}｜score={score:.2f}】\n{chunk}")
    return "\n\n".join(parts)


def list_knowledge_files() -> list[str]:
    return [name for name, _ in _read_kb_files()]
