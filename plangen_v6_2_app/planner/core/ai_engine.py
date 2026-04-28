from __future__ import annotations

import json
import requests

OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
OLLAMA_TAGS_URL = "http://127.0.0.1:11434/api/tags"
OLLAMA_MODEL = "qwen2.5:3b"

def _extract_json(text: str) -> dict:
    text = text.strip()
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start:end+1]
    try:
        return json.loads(text)
    except Exception as e:
        return {"_error": "JSON解析失败", "_detail": f"{e}\n\n原始返回前2000字符：\n{text[:2000]}"}

def check_ollama_status() -> tuple[bool, str]:
    try:
        resp = requests.get(OLLAMA_TAGS_URL, timeout=6)
        resp.raise_for_status()
        data = resp.json()
        names = []
        for m in data.get("models", []):
            name = m.get("name") or m.get("model") or ""
            if name:
                names.append(name)
        if not names:
            return False, "Ollama 已启动，但未检测到任何本地模型。请先执行：ollama pull qwen2.5:3b"
        if OLLAMA_MODEL in names:
            return True, f"本地AI状态正常：已连接 Ollama，检测到模型 {OLLAMA_MODEL}"
        return False, f"Ollama 已启动，但未找到模型 {OLLAMA_MODEL}。当前模型：{', '.join(names)}"
    except Exception as e:
        return False, f"无法连接到 Ollama（{e}）"

def call_ai_json(prompt: str, timeout: int = 60) -> dict:
    ok, msg = check_ollama_status()
    if not ok:
        return {"_error": "Ollama调用失败", "_detail": msg}
    try:
        resp = requests.post(
            OLLAMA_URL,
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=timeout
        )
        resp.raise_for_status()
        return _extract_json(resp.json().get("response", ""))
    except Exception as e:
        return {"_error": "Ollama调用失败", "_detail": str(e)}

def call_ai_text(prompt: str, timeout: int = 45) -> str:
    ok, msg = check_ollama_status()
    if not ok:
        return f"AI调用失败: {msg}"
    try:
        resp = requests.post(
            OLLAMA_URL,
            json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False},
            timeout=timeout
        )
        resp.raise_for_status()
        return resp.json().get("response", "").strip()
    except Exception as e:
        return f"AI调用失败: {e}"
