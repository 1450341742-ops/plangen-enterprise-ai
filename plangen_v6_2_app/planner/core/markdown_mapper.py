from __future__ import annotations

import re
from typing import Any, Dict, List, Tuple

PROCESS_KEYS = ["随机化要求", "随机要求", "复筛要求", "终止治疗要求", "退出研究要求", "禁止合并治疗要求", "限制用药要求", "药物剂量调整规则", "剂量调整", "关键时间窗要求"]


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
    text = re.sub(r"^[\-\*+·]\s*", "", text)
    return text.strip(" |\t")


def _extract_first(patterns: List[str], text: str, default: str = "") -> str:
    for p in patterns:
        m = re.search(p, text, re.S | re.I)
        if m:
            val = _strip_md(m.group(1))
            if val:
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
            if not raw.startswith("|"):
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
    for i, line in enumerate(lines):
        plain = _strip_md(line)
        if any(k in plain for k in start_keys):
            start = i + 1
            break
    if start is None:
        return ""
    end = len(lines)
    for j in range(start, len(lines)):
        raw = lines[j].strip()
        plain = _strip_md(raw)
        if stop_keys and any(k in plain for k in stop_keys) and (raw.startswith("#") or re.match(r"^\d+(\.\d+)+", plain)):
            end = j
            break
    return "\n".join(lines[start:end]).strip()


def _section_lines(md: str, start_keys: List[str], stop_keys: List[str]) -> List[str]:
    sec = _section(md, start_keys, stop_keys)
    lines = []
    for line in sec.splitlines():
        plain = _strip_md(line)
        if not plain or plain.startswith("|") or re.fullmatch(r"[-| ]+", plain):
            continue
        if plain in ["待填写", "xxxxxx", "xxxxx"]:
            continue
        lines.append(plain)
    return lines


def _find_label_value(md: str, labels: List[str]) -> str:
    for table in _parse_md_tables(md):
        for row in table:
            if len(row) >= 2:
                left = row[0].replace("：", "").replace(":", "").strip()
                if any(lab in left for lab in labels):
                    return _strip_md(row[1])
    for lab in labels:
        v = _extract_first([rf"{re.escape(lab)}[:：]\s*([^\n]+)"], md)
        if v:
            return v
    return ""


def _extract_title(md: str) -> str:
    project = _extract_first([r"###\s*1\.1\s*项目名称[:：]?\s*\n([^#]+?)\n\s*###", r"1\.1\s*项目名称[:：]?\s*\n([^#]+?)\n\s*1\.2", r"项目名称[:：]\s*([^\n]+)"], md)
    if project:
        return project
    for line in [l.strip() for l in md.splitlines() if l.strip()][:30]:
        if line.startswith("#") and "模式" not in line and "摘要" not in line:
            title = _strip_md(line)
            return re.sub(r"中心(稽查|质控)计划$", "", title).strip()
    return "待补充项目名称"


def _version_parts(v: str) -> Tuple[str, str]:
    v = _strip_md(v)
    if "/" in v:
        a, b = v.split("/", 1)
        return a.strip() or "V1.0", b.strip()
    m = re.search(r"(V\d+(?:\.\d+)*)\s*[/ ]?\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日)?", v, re.I)
    if m:
        return (m.group(1) or "V1.0"), (m.group(2) or "")
    return v or "V1.0", ""


def _extract_protocol_code(md: str) -> str:
    return _extract_first([r"方案编号[:：]\s*([A-Za-z0-9\-]+)", r"（方案编号[:：]\s*([A-Za-z0-9\-]+)）", r"\(([A-Z]{2,}[A-Z0-9\-]{3,})\)"], md)


def _is_process(text: str) -> bool:
    return any(k in text for k in PROCESS_KEYS)


def _extract_criteria_rows(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]], List[Dict[str, str]]]:
    rows = _table_rows_by_headers(md, ["方案", "重点关注"])
    criteria, exclusions, process = [], [], []
    for row in rows:
        left = _strip_md(row[0])
        right = _strip_md(row[1]) if len(row) > 1 else ""
        if not left or left in ["方案", "待补充"]:
            continue
        if _is_process(left):
            process.append({"requirement": left, "focus": right})
        elif left.startswith("排除标准") or "排除标准" in left[:10]:
            exclusions.append({"criterion": left, "ai_focus": right})
        elif left.startswith("入选标准") or "入选标准" in left[:10]:
            criteria.append({"criterion": left, "ai_focus": right})
        else:
            criteria.append({"criterion": left, "ai_focus": right})
    return criteria, exclusions, process


def _extract_assignment_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["序号", "质控流程"])
    return [{"seq": r[0], "process": r[1], "assignee": r[2] if len(r) > 2 else "", "plan_time": r[3] if len(r) > 3 else ""} for r in rows if r and r[0] != "序号"]


def _extract_risk_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["数据风险因素", "详细信息"])
    return [{"risk_factor": r[0], "detail": r[1] if len(r) > 1 else ""} for r in rows if r and r[0] != "数据风险因素"]


def _extract_subject_rows(md: str) -> List[Dict[str, str]]:
    rows = _table_rows_by_headers(md, ["筛选号", "基本情况"])
    return [{"subject_id": r[0], "summary": r[1] if len(r) > 1 else ""} for r in rows if r and r[0] != "筛选号"]


def _extract_endpoint_rows(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    rows = _table_rows_by_headers(md, ["目的", "终点"])
    primary, secondary = [], []
    for r in rows:
        if len(r) < 2:
            continue
        item = {"objective": r[0], "endpoint": r[1]}
        if "主要" in r[0] or "主要" in r[1]:
            primary.append(item)
        else:
            secondary.append(item)
    return primary, secondary


def _extract_safety_rows(md: str) -> List[Dict[str, str]]:
    section = _section(md, ["2.5.3.4", "安全性信息处理与报告"], ["2.5.3.5", "临床试验数据记录", "2.5.4"])
    rows = _table_rows_by_headers(section if section else md, ["类别", "重点关注"])
    clean_rows = []
    for r in rows:
        if not r or r[0] == "类别":
            continue
        focus = _strip_md(r[1] if len(r) > 1 else "")
        focus = re.sub(r"(\n?·\s*)+$", "", focus).strip()
        clean_rows.append({"category": r[0], "focus": focus})
    return clean_rows


def _extract_imp_spec_and_points(md: str) -> Tuple[str, str]:
    section = _section(md, ["2.5.4", "临床试验用药品管理"], ["2.5.5", "生物样本管理", "2.5.6"])
    rows = _table_rows_by_headers(section, ["试验药物规格"])
    spec = rows[0][0] if rows else ""
    lines = []
    for plain in _section_lines(md, ["2.5.4", "临床试验用药品管理"], ["2.5.5", "生物样本管理", "2.5.6"]):
        if "试验药物规格" in plain or plain == spec:
            continue
        if "质控范围" in plain or "质控分工" in plain or "序号 |" in plain:
            continue
        lines.append(plain)
    return spec, "\n".join(lines[:10])


def _extract_biosample_points(md: str) -> str:
    return "\n".join(_section_lines(md, ["2.5.5", "生物样本管理"], ["2.5.6", "中心实验室", "独立评估机构", "2.6", "三、"])[:10])


def _extract_lab_points(md: str) -> str:
    return "\n".join(_section_lines(md, ["2.5.6", "中心实验室及独立评估机构", "中心实验室"], ["2.6", "法规依据", "三、"])[:10])


def _extract_summary(md: str) -> str:
    sec = _section(md, ["摘要总结"], ["一、", "项目简介"])
    lines = []
    for line in sec.splitlines():
        plain = _strip_md(line)
        if plain and not plain.startswith("|") and not re.fullmatch(r"[-| ]+", plain):
            lines.append(plain)
    return "\n".join(lines[:18])


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

    audit_type = _extract_first([r"###\s*1\.2\s*质控类型[:：]?\s*\n([^#]+?)\n\s*###", r"1\.2\s*质控类型[:：]?\s*\n([^#]+?)\n\s*1\.3", r"质控类型[:：]\s*([^\n]+)"], md, "中心常规质控")
    sponsor2 = _extract_first([r"###\s*1\.3\s*申办方[:：]?\s*\n([^#]+?)\n\s*###", r"1\.3\s*申办方[:：]?\s*\n([^#]+?)\n\s*1\.4", r"申办方[:：]\s*([^\n|]+)"], md, sponsor)
    sponsor = _strip_md(sponsor2) or sponsor

    criteria, exclusions, process = _extract_criteria_rows(md)
    primary, secondary = _extract_endpoint_rows(md)
    imp_spec, imp_points = _extract_imp_spec_and_points(md)

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
        "BIOSAMPLE_DESCRIPTION": _extract_biosample_points(md),
        "LAB_DESCRIPTION": _extract_lab_points(md),
        "project": {"name": project_name, "sponsor": sponsor, "protocol_code": _extract_protocol_code(md), "version_no": version_no, "version_date": version_date, "audit_type": audit_type},
        "protocol_analysis": {"study_design": _extract_first([r"本临床试验为(.+?)(?:。|\n)", r"研究设计[:：]\s*([^\n]+)"], md, ""), "primary_endpoint": primary[0]["endpoint"] if primary else "", "key_criteria": "；".join([x["criterion"] for x in criteria[:4] + exclusions[:4]])},
        "criteria_ai_rows": criteria,
        "exclusion_ai_rows": exclusions,
        "process_requirement_rows": process,
        "primary_endpoint_rows": primary,
        "secondary_endpoint_rows": secondary,
        "risk_analysis_rows": _extract_risk_rows(md),
        "subject_rows": _extract_subject_rows(md),
        "assignment_rows": _extract_assignment_rows(md),
        "safety_focus_rows": _extract_safety_rows(md),
        "report_send_rows": [{"name": "待填写", "email": "待填写", "title_company": "待填写"}],
        "ai_audit_key_points": "\n".join([f"{i+1}. {r['criterion']}：{r['ai_focus']}" for i, r in enumerate(criteria[:5] + exclusions[:5])]),
        "rbqm_strategy": "",
        "interview_questions": "",
        "finding_capa_draft": "",
        "defect_rows": [],
    }
