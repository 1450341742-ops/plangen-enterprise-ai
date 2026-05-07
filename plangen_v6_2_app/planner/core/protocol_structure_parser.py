from __future__ import annotations

import re
from typing import Any, Dict, List

SECTION_PATTERNS = {
    "title": [r"项目名称[:：]\s*([^\n]+)", r"研究题目[:：]\s*([^\n]+)", r"试验题目[:：]\s*([^\n]+)"],
    "protocol_no": [r"方案编号[:：]\s*([^\n]+)", r"Protocol\s*No\.?[:：]?\s*([^\n]+)"],
    "sponsor": [r"申办方[:：]\s*([^\n]+)", r"申办者[:：]\s*([^\n]+)", r"Sponsor[:：]\s*([^\n]+)"],
    "phase": [r"(I{1,3}|Ⅲ|Ⅱ|Ⅰ|IV|Ⅳ)期临床试验", r"Phase\s*([I1-4V]+)"],
}


def _clean(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _first(patterns: List[str], text: str) -> str:
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            return m.group(1).strip().strip("；;。")
    return ""


def _section(text: str, starts: List[str], stops: List[str]) -> str:
    lines = text.splitlines()
    start = None
    for i, line in enumerate(lines):
        plain = line.strip()
        if any(k in plain for k in starts):
            start = i + 1
            break
    if start is None:
        return ""
    end = len(lines)
    for j in range(start, len(lines)):
        plain = lines[j].strip()
        if any(k in plain for k in stops):
            if re.match(r"^(\d+(\.\d+)*|[一二三四五六七八九十]+[、.])", plain) or len(plain) < 40:
                end = j
                break
    return "\n".join(lines[start:end]).strip()


def _extract_numbered_items(section_text: str, include_keywords: List[str]) -> List[str]:
    if not section_text:
        return []
    lines = [x.strip() for x in section_text.splitlines() if x.strip()]
    items: List[str] = []
    buf = ""
    for line in lines:
        is_new = bool(re.match(r"^((入选|排除)?标准)?\s*\d+[a-zA-Z]?([）).、.:：]|\s)", line)) or any(line.startswith(k) for k in include_keywords)
        if is_new and buf:
            items.append(buf.strip())
            buf = line
        else:
            buf = f"{buf} {line}".strip() if buf else line
    if buf:
        items.append(buf.strip())
    return items[:120]


def _focus_for_requirement(req: str) -> str:
    checks = []
    if any(k in req for k in ["年龄", "岁"]):
        checks.append("核对身份证/病历/知情同意签署日期，确认年龄范围符合方案要求")
    if any(k in req for k in ["知情同意", "ICF"]):
        checks.append("核对ICF版本、签署日期、签署人及伦理批准版本一致性，确认任何筛选程序前完成签署")
    if any(k in req for k in ["实验室", "血", "尿", "肌酐", "ALT", "AST", "胆红素", "血小板", "血红蛋白"]):
        checks.append("核对实验室报告、采样/报告时间窗、单位换算及EDC录入一致性")
    if any(k in req for k in ["用药", "治疗", "抗凝", "药物"]):
        checks.append("核对医嘱、用药记录、合并用药表及禁止/限制用药时间窗")
    if any(k in req for k in ["妊娠", "哺乳", "避孕"]):
        checks.append("核对妊娠检测、避孕告知/承诺及随访记录")
    if any(k in req for k in ["随机", "IWRS"]):
        checks.append("核对IWRS随机记录、随机时间、分层因素及首次给药/干预前后顺序")
    if not checks:
        checks.append("核对源文件、研究者判断依据、时间窗及EDC/源数据一致性，确认符合方案要求")
    return "；".join(checks)


def _extract_endpoints(text: str) -> List[Dict[str, str]]:
    sec = _section(text, ["研究目的", "研究终点", "主要终点", "次要终点"], ["入选标准", "排除标准", "研究设计", "安全性", "统计"])
    rows: List[Dict[str, str]] = []
    for line in [x.strip() for x in sec.splitlines() if x.strip()]:
        if any(k in line for k in ["主要目的", "主要终点"]):
            rows.append({"objective": "主要目的", "endpoint": line})
        elif any(k in line for k in ["次要目的", "次要终点"]):
            rows.append({"objective": "次要目的", "endpoint": line})
    return rows[:20]


def _extract_imp(text: str) -> str:
    sec = _section(text, ["试验用药品", "研究药物", "给药方案", "剂量", "用法用量"], ["生物样本", "实验室", "安全性", "统计", "入选标准"])
    if not sec:
        return ""
    lines = [x.strip() for x in sec.splitlines() if x.strip()]
    useful = [x for x in lines if any(k in x for k in ["规格", "剂型", "剂量", "给药", "用法", "调整", "药物", "试验药"])]
    return "；".join(useful[:12])


def _extract_sampling(text: str) -> str:
    return "优先抽取：首例/末例受试者、筛败/复筛病例、发生SAE/AESI或重要AE病例、主要终点事件病例、方案偏离病例、IMP剂量调整病例、实验室关键异常病例、生物样本采集/转运异常病例、关键时间窗临界病例；同时覆盖不同入组阶段和不同研究者执行记录。"


def parse_protocol_structure(text: str) -> Dict[str, Any]:
    text = _clean(text)
    incl_sec = _section(text, ["入选标准", "纳入标准", "Inclusion Criteria"], ["排除标准", "Exclusion Criteria", "随机", "研究设计", "用药"])
    excl_sec = _section(text, ["排除标准", "Exclusion Criteria"], ["随机", "研究设计", "给药", "研究终点", "安全性"])
    rand_sec = _section(text, ["随机", "IWRS", "分层因素"], ["给药", "研究终点", "安全性", "统计"])
    inclusion = _extract_numbered_items(incl_sec, ["入选标准"])
    exclusion = _extract_numbered_items(excl_sec, ["排除标准"])
    criteria_rows = [{"criterion": f"入选标准{i+1}：{x}", "ai_focus": _focus_for_requirement(x)} for i, x in enumerate(inclusion)]
    exclusion_rows = [{"criterion": f"排除标准{i+1}：{x}", "ai_focus": _focus_for_requirement(x)} for i, x in enumerate(exclusion)]
    process_rows = []
    if rand_sec:
        process_rows.append({"requirement": rand_sec[:1200], "focus": _focus_for_requirement(rand_sec)})
    endpoints = _extract_endpoints(text)
    imp = _extract_imp(text)
    return {
        "PROJECT_TITLE": _first(SECTION_PATTERNS["title"], text),
        "PROJECT_CODE": _first(SECTION_PATTERNS["protocol_no"], text),
        "SPONSOR_NAME": _first(SECTION_PATTERNS["sponsor"], text),
        "PHASE": _first(SECTION_PATTERNS["phase"], text),
        "SAMPLING_PRINCIPLE": _extract_sampling(text),
        "RISK_SAMPLING_RULE": _extract_sampling(text),
        "criteria_ai_rows": criteria_rows,
        "exclusion_ai_rows": exclusion_rows,
        "process_requirement_rows": process_rows,
        "primary_endpoint_rows": [r for r in endpoints if "主要" in r.get("objective", "")],
        "secondary_endpoint_rows": [r for r in endpoints if "主要" not in r.get("objective", "")],
        "IMP_SPEC": imp,
        "IMP_DESCRIPTION": imp,
        "LAW_SUPPLEMENT": "本项目质控应结合GCP、ICH E6(R3)、临床试验方案、研究者手册、伦理批准文件、试验用药品管理要求及数据完整性原则进行综合判断；重点关注受试者权益保护、方案依从性、源数据真实准确完整可溯源、关键时间窗执行及安全性信息报告。",
    }
