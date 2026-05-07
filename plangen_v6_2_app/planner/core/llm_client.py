from __future__ import annotations

import os
import requests
from typing import Any, Dict, List


class LLMClientError(Exception):
    pass


def build_plangen_prompt(user_prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
    return f"""
你是临床试验中心质控计划生成智能体。请根据用户输入和方案内容，输出可供PlanGen映射Word模板的标准Markdown。

必须输出以下模块，标题尽量保持一致：
1. 项目名称
2. 质控类型
3. 申办方
4. 摘要总结
5. 中心稽查风险评估病历抽取原则
6. 2.5.3.2 受试者筛选入组及方案执行：用表格输出，表头为：方案｜重点关注
7. 随机化/复筛/终止/退出/禁止合并治疗/剂量调整：用表格输出，表头为：方案描述｜重点关注
8. 2.5.3.3 研究目的和终点：用表格输出，表头为：主要目的｜主要终点
9. 2.5.4 临床试验用药品管理的审核：明确写出试验药物规格、剂型、剂量、给药方式、剂量调整
10. 2.6 法规依据补充说明

输出要求：
- 不要输出代码块。
- 不要输出闲聊内容。
- 不要省略入选标准和排除标准编号。
- 重点关注必须体现源文件、时间窗、证据链、EDC/源数据一致性。
- 如果无法从方案中确定字段，请写“待人工确认”，不要编造。

用户要求：
{user_prompt}

补充上下文：
{extra_context}

方案或原始资料：
{protocol_text}
""".strip()


def call_llm(user_prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
    """
    V8.2 通用AI客户端，兼容 OpenAI 格式服务。

    推荐配置：
    - 硅基流动：
      LLM_BASE_URL=https://api.siliconflow.cn/v1
      LLM_MODEL=deepseek-ai/DeepSeek-V3
    - DeepSeek官方：
      LLM_BASE_URL=https://api.deepseek.com
      LLM_MODEL=deepseek-chat
    - 阿里百炼 OpenAI兼容模式：按其控制台提供的 base_url/model 填写。

    必填环境变量：
    - LLM_API_KEY
    可选环境变量：
    - LLM_BASE_URL，默认 https://api.siliconflow.cn/v1
    - LLM_MODEL，默认 deepseek-ai/DeepSeek-V3
    """
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
            {"role": "system", "content": "你是临床试验质量管理和中心质控计划撰写专家，必须输出结构化Markdown。"},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.2,
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
    except Exception as e:
        raise LLMClientError(f"AI返回内容不是JSON：{resp.text[:800]}") from e

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
