from __future__ import annotations

import hashlib
import json
import math
import os
import pickle
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

import requests

BASE_DIR = Path(__file__).resolve().parents[2]
KB_DIR = BASE_DIR / "knowledge_base"
VECTOR_DIR = BASE_DIR / "vector_store"
VECTOR_DIR.mkdir(exist_ok=True)
INDEX_PATH = VECTOR_DIR / "kb_vector_index.pkl"
META_PATH = VECTOR_DIR / "kb_vector_meta.json"

DEFAULT_EMBED_BASE_URL = "https://api.siliconflow.cn/v1"
DEFAULT_EMBED_MODEL = "BAAI/bge-m3"


def _read_kb_files() -> list[tuple[str, str, str]]:
    if not KB_DIR.exists():
        return []
    items: list[tuple[str, str, str]] = []
    for ext in ("*.md", "*.txt"):
        for path in sorted(KB_DIR.glob(ext)):
            try:
                text = path.read_text(encoding="utf-8", errors="ignore")
                if text.strip():
                    items.append((path.name, str(path), text))
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


def _hash_files(files: list[tuple[str, str, str]]) -> str:
    h = hashlib.sha256()
    for name, path, text in files:
        h.update(name.encode("utf-8"))
        h.update(text.encode("utf-8"))
    return h.hexdigest()


def _fallback_embedding(text: str, dim: int = 384) -> list[float]:
    # 无embedding API时的本地兜底：哈希词袋向量，不需要额外依赖，适合免费部署。
    vec = [0.0] * dim
    tokens = re.findall(r"[a-zA-Z0-9_\-]+|[\u4e00-\u9fff]{1,2}", text.lower())
    for token in tokens:
        idx = int(hashlib.md5(token.encode("utf-8")).hexdigest(), 16) % dim
        vec[idx] += 1.0
    return _normalize(vec)


def _normalize(vec: list[float]) -> list[float]:
    norm = math.sqrt(sum(x * x for x in vec))
    if norm == 0:
        return vec
    return [x / norm for x in vec]


def _embed_remote(texts: list[str]) -> list[list[float]]:
    api_key = os.getenv("EMBEDDING_API_KEY", os.getenv("LLM_API_KEY", "")).strip()
    base_url = os.getenv("EMBEDDING_BASE_URL", os.getenv("LLM_BASE_URL", DEFAULT_EMBED_BASE_URL)).strip().rstrip("/")
    model = os.getenv("EMBEDDING_MODEL", DEFAULT_EMBED_MODEL).strip()

    if not api_key:
        return [_fallback_embedding(t) for t in texts]

    url = f"{base_url}/embeddings"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload: Dict[str, Any] = {"model": model, "input": texts}

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=120)
        if resp.status_code >= 400:
            return [_fallback_embedding(t) for t in texts]
        data = resp.json()
        embeddings = []
        for item in data.get("data", []):
            emb = item.get("embedding")
            if isinstance(emb, list):
                embeddings.append(_normalize([float(x) for x in emb]))
        if len(embeddings) == len(texts):
            return embeddings
    except Exception:
        pass

    return [_fallback_embedding(t) for t in texts]


def build_vector_index(force: bool = False) -> Dict[str, Any]:
    files = _read_kb_files()
    file_hash = _hash_files(files)

    if not force and INDEX_PATH.exists() and META_PATH.exists():
        try:
            meta = json.loads(META_PATH.read_text(encoding="utf-8"))
            if meta.get("file_hash") == file_hash:
                with INDEX_PATH.open("rb") as f:
                    return pickle.load(f)
        except Exception:
            pass

    chunks: list[dict[str, Any]] = []
    texts: list[str] = []
    for filename, path, text in files:
        for i, chunk in enumerate(_split_chunks(text)):
            chunks.append({"file": filename, "chunk_id": i, "text": chunk})
            texts.append(chunk)

    embeddings: list[list[float]] = []
    batch_size = 16
    for i in range(0, len(texts), batch_size):
        embeddings.extend(_embed_remote(texts[i:i + batch_size]))

    index = {"chunks": chunks, "embeddings": embeddings, "file_hash": file_hash}
    with INDEX_PATH.open("wb") as f:
        pickle.dump(index, f)
    META_PATH.write_text(json.dumps({"file_hash": file_hash, "chunk_count": len(chunks)}, ensure_ascii=False, indent=2), encoding="utf-8")
    return index


def _cosine(a: list[float], b: list[float]) -> float:
    if not a or not b or len(a) != len(b):
        return 0.0
    return sum(x * y for x, y in zip(a, b))


def retrieve_vector_knowledge(query: str, top_k: int = 8) -> str:
    index = build_vector_index(force=False)
    chunks = index.get("chunks", [])
    embeddings = index.get("embeddings", [])
    if not chunks or not embeddings:
        return ""
    q_emb = _embed_remote([query])[0]
    scored: list[tuple[float, dict[str, Any]]] = []
    for chunk, emb in zip(chunks, embeddings):
        scored.append((_cosine(q_emb, emb), chunk))
    scored.sort(key=lambda x: x[0], reverse=True)
    parts = []
    for i, (score, chunk) in enumerate(scored[:top_k], 1):
        parts.append(f"【向量知识片段{i}｜{chunk['file']}｜score={score:.3f}】\n{chunk['text']}")
    return "\n\n".join(parts)


def vector_index_status() -> Dict[str, Any]:
    if META_PATH.exists():
        try:
            return json.loads(META_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"file_hash": "", "chunk_count": 0}
