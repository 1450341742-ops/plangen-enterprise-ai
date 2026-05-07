from __future__ import annotations

import os
from pathlib import Path
import requests
from typing import Any, Dict, List

from planner.core.rag_engine import retrieve_knowledge, list_knowledge_files


class LLMClientError(Exception):
    pass


BASE_DIR = Path(__file__).resolve().parents[2]
KB_DIR = BASE_DIR / "knowledge_base"


def load_knowledge_base() -> str:
    if not KB_DIR.exists():
        return ""

    parts: List[str] = []

    for path in sorted(KB_DIR.glob("*.md")):
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
            if text.strip():
                parts.append(f"\n\n# 知识库文件：{path.name}\n{text}")
        except Exception:
            continue

    for path in sorted(KB_DIR.glob("*.txt")):
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
            if text.strip():
                parts.append(f"\n\n# 知识库文件：{path.name}\n{text}")
        except Exception:
            continue

    return "\n".join(parts)


def build_plangen_prompt(user_prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
    kb_text = load_knowledge_base()

    rag_query = f"{user_prompt}\n{protocol_text[:4000]}"
    rag_context = retrieve_knowledge(rag_query, top_k=8)
    kb_files = list_knowledge_files()

    return f"""
你是万宁睿和第三方稽查公司AI质控专家。
请严格参考企业知识库、RAG召回内容、RBQM规则、FDA/CFDI核查逻辑生成中心质控计划。

# 系统要求
- 输出必须可供PlanGen映射Word模板。
- 必须输出结构化Markdown。
- 不允许输出闲聊。
- 不允许输出代码块。
- 不允许编造方案中不存在的数据。
- 如无法确认，请写“待人工确认”。

# 必须输出模块
1. 项目名称
2. 质控类型
3. 申办方
4. 摘要总结
5. 中心稽查风险评估病历抽取原则
6. 2.5.3.2 受试者筛选入组及方案执行
7. 随机化/复筛/终止/退出/禁止合并治疗/剂量调整
8. 2.5.3.3 研究目的和终点
9. 2.5.4 临床试验用药品管理的审核
10. 2.5.5 生物样本管理
11. 2.5.6 中心实验室及独立评估机构
12. 2.6 法规依据补充说明

# 输出要求
- 不要省略入选标准和排除标准编号。
- 重点关注必须体现源文件、时间窗、证据链、EDC/源数据一致性。
- IMP部分必须明确规格、剂型、剂量、给药方式、剂量调整。
- 生物样本管理必须包含：采集、运输、保存、时间窗、异常处理。
- 中心实验室部分必须包含：CAP/CLIA、EDC、IWRS、供应商管理。
- 输出风格应接近资深临床稽查经理。

# 用户要求
{user_prompt}

# 补充上下文
{extra_context}

# 方案或原始资料
{protocol_text}

# 企业知识库文件
{', '.join(kb_files)}

# RAG召回知识片段
{rag_context}

# 企业完整知识库
{kb_text}
""".strip()


def call_llm(user_prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
    api_key = os.getenv("LLM_API_KEY", "").strip()
    base_url = os.getenv("LLM_BASE_URL", "https://api.siliconflow.cn/v1").strip().rstrip("/")
    model = os.getenv("LLM_MODEL", "deepseek-ai/DeepSeek-V3").strip()

    if not api_key:
        raise LLMClientError("未配置 LLM_API_KEY，请在部署环境变量或 Streamlit Secrets 中填写。")

    prompt = build_plangen_prompt(user_prompt=user_prompt, protocol_text=protocol_text, extra_context=extra_context)

    url = f"{base_url}/chat/completions"

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    payload: Dict[str, Any] = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "你是临床试验质量管理、RBQM、中心稽查、FDA/CFDI核查专家。",
            },
            {
                "role": "user",
                "content": prompt,
            },
        ],
        "temperature": 0.1,
        "stream": False,
    }

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=240)
    except Exception as e:
        raise LLMClientError(f"调用通用AI失败：{e}") from e

    if resp.status_code >= 400:
        raise LLMClientError(f"通用AI接口返回错误：{resp.status_code} {resp.text[:800]}")

    try:
        data = resp.json()
    except Exception:
        raise LLMClientError(f"AI返回内容不是JSON：{resp.text[:800]}")

    choices: List[Dict[str, Any]] = data.get("choices", [])

    if choices:
        msg = choices[0].get("message", {})
        content = msg.get("content")
        if isinstance(content, str) and content.strip():
            return content

    for key in ["content", "answer", "result", "output", "text"]:
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            return value

    raise LLMClientError(f"未从AI返回中解析到文本内容：{str(data)[:800]}")
