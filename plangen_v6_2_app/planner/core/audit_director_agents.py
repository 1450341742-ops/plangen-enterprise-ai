from __future__ import annotations

from typing import Dict

from planner.core.risk_profile_engine import build_risk_profile
from planner.core.capa_learning import extract_capa_patterns
from planner.core.project_memory import retrieve_project_memory
from planner.core.vector_store import retrieve_vector_knowledge


def rbqm_agent(protocol_text: str) -> str:
    risk = build_risk_profile(protocol_text)
    return "\n".join([
        "# AI RBQM专家意见",
        f"风险画像：{risk.get('risk_summary', '')}",
        "建议优先关注：入排标准、主要终点、安全性事件、IMP管理、生物样本、中心实验室、关键时间窗、EDC/源数据一致性。",
    ])


def fda_cfdi_agent(protocol_text: str) -> str:
    return "\n".join([
        "# AI FDA/CFDI核查专家意见",
        "核查重点应围绕受试者权益保护、知情同意过程、方案依从性、主要疗效终点、安全性报告、数据可靠性与可溯源性展开。",
        "建议所有重点关注均落到源文件、系统截图、时间窗、版本一致性、研究者判断依据。",
    ])


def capa_agent(protocol_text: str) -> str:
    capa = extract_capa_patterns(protocol_text)
    patterns = capa.get("patterns", [])[:6]
    lines = ["# AI CAPA专家意见"]
    if patterns:
        for p in patterns:
            lines.append(f"- {p['category']}：{p['recommendation']}")
    else:
        lines.append("- 暂未识别明确CAPA模式，建议基于稽查发现后再形成CAPA闭环。")
    return "\n".join(lines)


def project_manager_agent(project_name: str, protocol_text: str) -> str:
    memory = retrieve_project_memory(f"{project_name}\n{protocol_text[:2000]}", top_k=3)
    return "\n".join([
        "# AI项目经理意见",
        "建议本项目生成前先核对项目名称、方案编号、申办方、版本号/版本日期、质控类型、模板版本。",
        "建议输出后人工复核：入排完整性、IMP规格、法规依据、模板请填写区域是否全部替换。",
        memory if memory else "暂无相似历史项目记忆。",
    ])


def audit_director_brief(project_name: str, protocol_text: str, user_prompt: str = "") -> Dict[str, str]:
    query = f"{project_name}\n{user_prompt}\n{protocol_text[:3000]}"
    rag = retrieve_vector_knowledge(query, top_k=6)
    return {
        "project_manager": project_manager_agent(project_name, protocol_text),
        "rbqm": rbqm_agent(protocol_text),
        "fda_cfdi": fda_cfdi_agent(protocol_text),
        "capa": capa_agent(protocol_text),
        "rag": rag,
        "combined": "\n\n".join([
            project_manager_agent(project_name, protocol_text),
            rbqm_agent(protocol_text),
            fda_cfdi_agent(protocol_text),
            capa_agent(protocol_text),
            "# 企业向量知识召回",
            rag,
        ]),
    }
