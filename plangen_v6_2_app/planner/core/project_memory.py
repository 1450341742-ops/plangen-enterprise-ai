from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

BASE_DIR = Path(__file__).resolve().parents[2]
MEMORY_DIR = BASE_DIR / "memory_store"
MEMORY_DIR.mkdir(exist_ok=True)
MEMORY_PATH = MEMORY_DIR / "project_memory.jsonl"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _safe_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _tokenize(text: str) -> set[str]:
    return set(re.findall(r"[a-zA-Z0-9_\-]+|[\u4e00-\u9fff]{1,2}", text.lower()))


def save_project_memory(project_name: str, source_text: str, generated_markdown: str = "", metadata: Dict[str, Any] | None = None) -> None:
    record = {
        "created_at": _now(),
        "project_name": _safe_text(project_name) or "未命名项目",
        "source_text": _safe_text(source_text)[:12000],
        "generated_markdown": _safe_text(generated_markdown)[:12000],
        "metadata": metadata or {},
    }
    with MEMORY_PATH.open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def load_project_memories(limit: int = 200) -> List[Dict[str, Any]]:
    if not MEMORY_PATH.exists():
        return []
    rows: List[Dict[str, Any]] = []
    for line in MEMORY_PATH.read_text(encoding="utf-8", errors="ignore").splitlines():
        try:
            rows.append(json.loads(line))
        except Exception:
            continue
    return rows[-limit:]


def retrieve_project_memory(query: str, top_k: int = 5) -> str:
    memories = load_project_memories(limit=300)
    if not memories:
        return ""
    q = _tokenize(query)
    scored = []
    for m in memories:
        text = "\n".join([
            m.get("project_name", ""),
            m.get("source_text", ""),
            m.get("generated_markdown", ""),
        ])
        tokens = _tokenize(text)
        score = len(q & tokens)
        if score > 0:
            scored.append((score, m))
    scored.sort(key=lambda x: x[0], reverse=True)
    parts = []
    for i, (score, m) in enumerate(scored[:top_k], 1):
        parts.append(
            f"【历史项目记忆{i}｜{m.get('project_name')}｜{m.get('created_at')}｜score={score}】\n"
            f"生成摘要：\n{m.get('generated_markdown','')[:1800]}"
        )
    return "\n\n".join(parts)


def memory_status() -> Dict[str, Any]:
    memories = load_project_memories(limit=100000)
    return {
        "memory_count": len(memories),
        "latest_project": memories[-1].get("project_name") if memories else "",
        "latest_time": memories[-1].get("created_at") if memories else "",
    }
