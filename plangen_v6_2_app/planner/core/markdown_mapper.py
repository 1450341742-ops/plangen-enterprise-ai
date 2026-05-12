from __future__ import annotations

import re
from typing import Any, Dict, List, Tuple

PROCESS_KEYS = ["随机化要求", "随机要求", "复筛要求", "终止治疗要求", "退出研究要求", "禁止合并治疗要求", "限制用药要求", "药物剂量调整规则", "剂量调整", "关键时间窗要求", "随机化程序"]
PLACEHOLDERS = {"待填写", "xxxxx", "xxxxxx", "xxx", "[填写]", "其他", "请填写"}


def _clean(text: str) -> str:
    if text is None:
        return ""
    text = str(text).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.I)
    text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
    text = re.sub(r"`([^`]*)`", r"\1", text)
    text = text.replace("[填写]", "待填写").replace("&nbsp;", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def _strip_md(text: str) -> str:
    text = _clean(text)
    text = re.sub(r"^#{1,6}\s*", "", text)
    text = re.sub(r"^[\-\*+·■▪•]\s*", "", text)
    return text.strip(" |\t")


def _valid(text: str) -> bool:
    t = _strip_md(text)
    return bool(t) and t not in PLACEHOLDERS and not re.fullmatch(r"[-| ]+", t)


def _extract_first(patterns: List[str], text: str, default: str = "") -> str:
    for p in patterns:
        m = re.search(p, text, re.S | re.I)
        if m:
            val = _strip_md(m.group(1))
            if _valid(val):
                return val
    return default


def _parse_md_tables(md: str) -> List[List[List[str]]]:
    tables = []
    lines = md.splitlines()
    i = 0
    while i < len(lines):
        if "|" not in lines[i]:
            i += 1
            continue
        block = []
        while i < len(lines) and "|" in lines[i]:
            block.append(lines[i])
            i += 1
        rows = []
        for b in block:
            raw = b.strip()
            if not raw.startswith("|") and raw.count("|") >= 1:
                raw = "|" + raw + "|"
            if "|" not in raw:
                continue
            cells = [_strip_md(c) for c in raw.strip("|").split("|")]
            if cells and not all(re.fullmatch(r":?-{2,}:?", c.replace(" ", "")) for c in cells):
                rows.append(cells)
        if rows:
            tables.append(rows)
    return tables


def _table_rows_by_headers(md: str, header_keywords: List[str]) -> List[List[str]]:
    out = []
    for table in _parse_md_tables(md):
        header = "|".join(table[0]) if table else ""
        if all(k in header for k in header_keywords):
            out.extend([r for r in table[1:] if len(r) >= 2 and any(c.strip() for c in r)])
    return out


def _section(md: str, start_keys: List[str], stop_keys: List[str] | None = None) -> str:
    lines = md.splitlines()
    start = None
    prefix = ""
    for i, line in enumerate(lines):
        plain = _strip_md(line)
        if any(k in plain for k in start_keys):
            start = i + 1
            if "：" in plain or ":" in plain:
                parts = re.split(r"[:：]", plain, maxsplit=1)
                if len(parts) > 1 and _valid(parts[1]):
                    prefix = parts[1].strip()
            break
    if start is None:
        return ""
    end = len(lines)
    for j in range(start, len(lines)):
        plain = _strip_md(lines[j])
        if not plain:
            continue
        if stop_keys and any(k in plain for k in stop_keys):
            end = j
            break
    body = "\n".join(lines[start:end]).strip()
    return (prefix + "\n" + body).strip() if prefix else body


def _find_label_value(md: str, labels: List[str]) -> str:
    for table in _parse_md_tables(md):
        for row in table:
            if len(row) >= 2:
                left = row[0].replace("：", "").replace(":", "").strip()
                if any(lab in left for lab in labels):
                    val = _strip_md(row[1])
                    if _valid(val):
                        return val
    for lab in labels:
        v = _extract_first([rf"{re.escape(lab)}[:：]\s*([^\n]+)"], md)
        if v:
            return v
    return ""


def _extract_title(md: str) -> str:
    project = _extract_first([r"项目名称[:：]\s*([^\n]+)", r"1\.1\s*项目名称[:：]?\s*\n([^#\n]+)"], md)
    if project:
        return project
    for line in [l.strip() for l in md.splitlines() if l.strip()][:30]:
        if line.startswith("#") and "摘要" not in line and "模式" not in line:
            title = _strip_md(line)
            if _valid(title):
                return re.sub(r"中心(稽查|质控)计划$", "", title).strip()
    return "待补充项目名称"


def _version_parts(v: str) -> Tuple[str, str]:
    v = _strip_md(v)
    if "/" in v:
        a, b = v.split("/", 1)
        return a.strip() or "V1.0", b.strip()
    return v or "V1.0", ""


def _extract_protocol_code(md: str) -> str:
    return _extract_first([r"方案编号[:：]\s*([A-Za-z0-9\-]+)", r"\(([A-Z]{2,}[A-Z0-9\-]{3,})\)"], md)


def _is_process(text: str) -> bool:
    return any(k in text for k in PROCESS_KEYS)


def _make_focus_from_criterion(text: str) -> str:
    checks = []
    if any(k in text for k in ["年龄", "岁"]):
        checks.append("核对身份证/病历/ICF签署日期，确认年龄范围符合方案要求")
    if any(k in text for k in ["知情同意", "ICF"]):
        checks.append("核对ICF版本、签署日期、签署人及筛选程序前完成情况")
    if any(k in text for k in ["实验室", "血", "尿", "肌酐", "ALT", "AST", "胆红素", "血小板", "血红蛋白", "纤维蛋白原"]):
        checks.append("核对实验室报告、采样时间窗、单位换算及EDC/源数据一致性")
    if any(k in text for k in ["用药", "药物", "华法林", "阿司匹林", "肝素"]):
        checks.append("核对医嘱、合并用药记录及禁止/限制用药时间窗")
    if any(k in text for k in ["妊娠", "哺乳", "避孕"]):
        checks.append("核对妊娠检测、避孕告知/承诺及随访记录")
    if not checks:
        checks.append("核对源文件、研究者判断依据、时间窗及EDC/源数据一致性")
    return "；".join(checks)


def _extract_criteria_rows(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]], List[Dict[str, str]]]:
    rows = _table_rows_by_headers(md, ["方案", "重点关注"])
    criteria, exclusions, process = [], [], []
    for row in rows:
        left = _strip_md(row[0])
        right = _strip_md(row[1]) if len(row) > 1 else ""
        if not _valid(left) or left in ["方案", "待补充", "请填写"]:
            continue
        if _is_process(left):
            process.append({"requirement": left, "focus": right})
        elif left.startswith("排除标准") or "排除标准" in left[:12]:
            exclusions.append({"criterion": left, "ai_focus": right or _make_focus_from_criterion(left)})
        elif left.startswith("入选标准") or "入选标准" in left[:12]:
            criteria.append({"criterion": left, "ai_focus": right or _make_focus_from_criterion(left)})
        else:
            criteria.append({"criterion": left, "ai_focus": right or _make_focus_from_criterion(left)})

    # 兜底：钉钉未输出表格时，从“入选标准/排除标准”文本块逐条抽取
    if not criteria and not exclusions:
        incl = _section(md, ["入选标准", "纳入标准"], ["排除标准", "随机化", "随机", "研究目的", "研究终点", "2.5.3.3"])
        excl = _section(md, ["排除标准"], ["随机化", "随机", "研究目的", "研究终点", "2.5.3.3", "2.5.4"])
        for line in incl.splitlines():
            t = _strip_md(line)
            if _valid(t) and (re.match(r"^\d+", t) or "入选" in t):
                item = t if t.startswith("入选标准") else f"入选标准：{t}"
                criteria.append({"criterion": item, "ai_focus": _make_focus_from_criterion(item)})
        for line in excl.splitlines():
            t = _strip_md(line)
            if _valid(t) and (re.match(r"^\d+", t) or "排除" in t):
                item = t if t.startswith("排除标准") else f"排除标准：{t}"
                exclusions.append({"criterion": item, "ai_focus": _make_focus_from_criterion(item)})
    return criteria, exclusions, process


def _extract_endpoint_rows(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    rows = _table_rows_by_headers(md, ["目的", "终点"])
    primary, secondary = [], []
    for r in rows:
        if len(r) < 2 or not _valid(r[0]) or not _valid(r[1]):
            continue
        item = {"objective": r[0], "endpoint": r[1]}
        if "主要" in r[0] or "主要" in r[1]:
            primary.append(item)
        else:
            secondary.append(item)
    if not primary and not secondary:
        sec = _section(md, ["研究目的和终点", "研究目的", "研究终点", "主要目的", "主要终点"], ["安全性信息", "2.5.3.4", "2.5.4", "临床试验用药品"])
        objective = _extract_first([r"主要目的[:：]?\s*([^\n]+)", r"目的[:：]?\s*([^\n]+)"], sec)
        endpoint = _extract_first([r"主要终点[:：]?\s*([^\n]+)"], sec)
        if objective or endpoint:
            primary.append({"objective": objective or "主要目的", "endpoint": endpoint or "主要终点待确认"})
        for line in sec.splitlines():
            t = _strip_md(line)
            if _valid(t) and "次要终点" in t:
                secondary.append({"objective": "次要目的", "endpoint": t})
    return primary, secondary


def _extract_imp_spec_and_points(md: str) -> Tuple[str, str]:
    section = _section(md, ["2.5.4", "临床试验用药品管理"], ["2.5.5", "生物样本管理", "2.5.6", "三、"])
    spec = _extract_first([r"试验药物规格[^\n]*[:：]\s*([^\n]+)", r"规格[^\n]*[:：]\s*([^\n]+)"], section)
    return spec, ""


def _extract_summary(md: str) -> str:
    sec = _section(md, ["摘要总结"], ["一、", "项目简介"])
    lines = [_strip_md(x) for x in sec.splitlines() if _valid(x)]
    return "\n".join(lines[:18])


def _extract_sampling_principle(md: str) -> str:
    sec = _section(md, ["中心稽查风险评估病历抽取原则", "中心质控风险评估病历抽取原则", "抽取原则"], ["中心筛选情况", "2.5.1", "2.5.2", "2.5.3", "受试者筛选", "三、"])
    lines = []
    for line in sec.splitlines():
        plain = _strip_md(line)
        if _valid(plain) and "抽取原则" not in plain:
            lines.append(plain)
    return "\n".join(lines[:12])


def _extract_law_supplement(md: str) -> str:
    sec = _section(md, ["2.6", "法规依据补充说明"], ["三、", "3.1", "质控流程"])
    lines = [_strip_md(x) for x in sec.splitlines() if _valid(x) and "2.6" not in _strip_md(x)]
    return "\n".join(lines[:12]) or "本项目质控应结合现行GCP、临床试验方案、研究者手册、伦理批准文件及相关指导原则进行综合判断，确保受试者权益、安全保护及数据真实、准确、完整、可溯源。"


def _extract_risk_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["数据风险因素", "详细信息"])
    return [{"risk_factor": r[0], "detail": r[1] if len(r) > 1 else ""} for r in rows if r and r[0] != "数据风险因素"]


def _extract_subject_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["筛选号", "基本情况"])
    return [{"subject_id": r[0], "summary": r[1] if len(r) > 1 else ""} for r in rows if r and r[0] != "筛选号"]


def _extract_assignment_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["序号", "质控流程"])
    return [{"seq": r[0], "process": r[1], "assignee": r[2] if len(r) > 2 else "", "plan_time": r[3] if len(r) > 3 else ""} for r in rows if r and r[0] != "序号"]


def parse_markdown_to_template_data(md_text: str) -> Dict[str, Any]:
    md = md_text.replace("\r\n", "\n")
    project_name = _extract_title(md)
    sponsor = _find_label_value(md, ["申办方", "申办者"])
    company = _find_label_value(md, ["质控公司", "稽查公司"]) or "北京万宁睿和医药科技有限公司"
    version_no, version_date = _version_parts(_find_label_value(md, ["版本号/版本日期", "版本号", "版本日期"]))
    author_approver = _find_label_value(md, ["撰写人/审批人", "撰写人"])
    if "/" in author_approver:
        author, approver = [x.strip() for x in author_approver.split("/", 1)]
    else:
        author, approver = author_approver or "苗田", "待审批"
    audit_type = _extract_first([r"1\.2\s*质控类型[:：]?\s*\n([^#\n]+)", r"质控类型[:：]\s*([^\n]+)"], md, "中心常规质控")
    criteria, exclusions, process = _extract_criteria_rows(md)
    primary, secondary = _extract_endpoint_rows(md)
    imp_spec, imp_points = _extract_imp_spec_and_points(md)
    sampling = _extract_sampling_principle(md)
    return {
        "AUTHOR": author,
        "APPROVER": approver,
        "AUDIT_COMPANY": company,
        "PROJECT_TITLE": project_name,
        "PROJECT_CODE": _extract_protocol_code(md),
        "SPONSOR_NAME": sponsor,
        "VERSION_NO": version_no,
        "VERSION_DATE": version_date,
        "AUDIT_TYPE": audit_type,
        "SUMMARY_TEXT": _extract_summary(md),
        "IMP_SPEC": imp_spec,
        "IMP_DESCRIPTION": imp_points,
        "SAMPLING_PRINCIPLE": sampling,
        "RISK_SAMPLING_RULE": sampling,
        "LAW_SUPPLEMENT": _extract_law_supplement(md),
        "project": {"name": project_name, "sponsor": sponsor, "protocol_code": _extract_protocol_code(md), "version_no": version_no, "version_date": version_date, "audit_type": audit_type},
        "criteria_ai_rows": criteria,
        "exclusion_ai_rows": exclusions,
        "process_requirement_rows": process,
        "primary_endpoint_rows": primary,
        "secondary_endpoint_rows": secondary,
        "risk_analysis_rows": _extract_risk_rows(md),
        "subject_rows": _extract_subject_rows(md),
        "assignment_rows": _extract_assignment_rows(md),
        "report_send_rows": [{"name": "待填写", "email": "待填写", "title_company": "待填写"}],
        "defect_rows": [],
    }
