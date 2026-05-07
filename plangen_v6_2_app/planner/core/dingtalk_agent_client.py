from __future__ import annotations

import os
import requests
from typing import Any, Dict


class DingTalkAgentError(Exception):
    pass


def call_dingtalk_agent(prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
    """
    通用钉钉智能体调用客户端。

    说明：
    1. 不同企业钉钉智能体开放接口字段可能不同，因此这里采用环境变量配置方式。
    2. 默认按 HTTP POST JSON 调用。
    3. 返回值统一转为文本，后续交给 PlanGen markdown_mapper 解析。

    必需环境变量：
    - DINGTALK_AGENT_API_URL：钉钉智能体接口地址

    可选环境变量：
    - DINGTALK_AGENT_API_KEY：如智能体接口需要鉴权，放到 Authorization Bearer 中
    - DINGTALK_AGENT_APP_ID：智能体/应用ID
    - DINGTALK_AGENT_USER_ID：默认调用用户ID
    """
    api_url = os.getenv("DINGTALK_AGENT_API_URL", "").strip()
    api_key = os.getenv("DINGTALK_AGENT_API_KEY", "").strip()
    app_id = os.getenv("DINGTALK_AGENT_APP_ID", "").strip()
    user_id = os.getenv("DINGTALK_AGENT_USER_ID", "plangen-system").strip()

    if not api_url:
        raise DingTalkAgentError("未配置 DINGTALK_AGENT_API_URL")

    final_prompt = build_agent_prompt(prompt=prompt, protocol_text=protocol_text, extra_context=extra_context)

    payload: Dict[str, Any] = {
        "app_id": app_id,
        "user_id": user_id,
        "prompt": final_prompt,
        "input": final_prompt,
        "query": final_prompt,
    }

    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"

    try:
        resp = requests.post(api_url, json=payload, headers=headers, timeout=180)
    except Exception as e:
        raise DingTalkAgentError(f"调用钉钉智能体失败：{e}") from e

    if resp.status_code >= 400:
        raise DingTalkAgentError(f"钉钉智能体接口返回错误：{resp.status_code} {resp.text[:500]}")

    try:
        data = resp.json()
    except Exception:
        return resp.text

    # 兼容常见返回字段
    for key in ["markdown", "content", "answer", "result", "output", "text"]:
        value = data.get(key)
        if isinstance(value, str) and value.strip():
            return value

    # 兼容嵌套结构
    if isinstance(data.get("data"), dict):
        inner = data["data"]
        for key in ["markdown", "content", "answer", "result", "output", "text"]:
            value = inner.get(key)
            if isinstance(value, str) and value.strip():
                return value

    return str(data)


def build_agent_prompt(prompt: str, protocol_text: str = "", extra_context: str = "") -> str:
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

注意：
- 不要输出与模板固定文本无关的闲聊内容。
- 不要输出代码块。
- 不要省略入选标准和排除标准编号。
- 重点关注必须可用于稽查执行，体现源文件、时间窗、证据链、EDC/源数据一致性。

用户要求：
{prompt}

补充上下文：
{extra_context}

方案或原始资料：
{protocol_text}
""".strip()
