from __future__ import annotations

import re
from typing import Any, Dict, List, Tuple

PLACEHOLDERS = {"待填写", "xxxxx", "xxxxxx", "xxx", "[填写]", "请填写", "其他"}
PROCESS_KEYS = ["随机化", "随机程序", "复筛", "终止治疗", "退出研究", "禁止合并", "限制用药", "剂量调整", "关键时间窗", "盲态", "揭盲"]


def _clean(s: Any) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    s = re.sub(r"\*\*(.*?)\*\*", r"\1", s)
    s = re.sub(r"`([^`]*)`", r"\1", s)
    return s.replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&").replace("&nbsp;", " ").strip()


def _strip(s: Any) -> str:
    s = _clean(s)
    s = re.sub(r"^#{1,6}\s*", "", s)
    s = re.sub(r"^[\-\*+·■▪•]\s*", "", s)
    return s.strip(" |\t")


def _ok(s: Any) -> bool:
    t = _strip(s)
    return bool(t) and t not in PLACEHOLDERS and not re.fullmatch(r"[-| ]+", t)


def _first(patterns: List[str], text: str, default: str = "") -> str:
    for p in patterns:
        m = re.search(p, text, re.S | re.I)
        if m:
            v = _strip(m.group(1))
            if _ok(v):
                return v
    return default


def _section(md: str, starts: List[str], stops: List[str] | None = None) -> str:
    lines = md.splitlines()
    start = None
    for i, line in enumerate(lines):
        if any(k in _strip(line) for k in starts):
            start = i + 1
            break
    if start is None:
        return ""
    end = len(lines)
    for j in range(start, len(lines)):
        plain = _strip(lines[j])
        if stops and plain and any(k in plain for k in stops):
            end = j
            break
    return "\n".join(lines[start:end]).strip()


def _tables(md: str) -> List[List[List[str]]]:
    out: List[List[List[str]]] = []
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
        for raw in block:
            raw = raw.strip()
            if not raw.startswith("|"):
                raw = "|" + raw + "|"
            cells = [_strip(c) for c in raw.strip("|").split("|")]
            if cells and not all(re.fullmatch(r":?-{2,}:?", c.replace(" ", "")) for c in cells):
                rows.append(cells)
        if rows:
            out.append(rows)
    return out


def _rows(md: str, headers: List[str]) -> List[List[str]]:
    res: List[List[str]] = []
    for t in _tables(md):
        head = "|".join(t[0]) if t else ""
        if all(h in head for h in headers):
            res.extend([r for r in t[1:] if len(r) >= 2 and any(_ok(c) for c in r)])
    return res


def _label(md: str, labels: List[str]) -> str:
    for t in _tables(md):
        for r in t:
            if len(r) >= 2 and any(x in r[0] for x in labels) and _ok(r[1]):
                return _strip(r[1])
    for lab in labels:
        v = _first([rf"{re.escape(lab)}[:：]\s*([^\n]+)"], md)
        if v:
            return v
    return ""


def _title(md: str) -> str:
    return _first([r"项目名称[:：]\s*([^\n]+)", r"1\.1\s*项目名称[:：]?\s*\n([^#\n]+)"], md, "待补充项目名称")


def _version(v: str) -> Tuple[str, str]:
    v = _strip(v)
    if "/" in v:
        a, b = v.split("/", 1)
        return a.strip() or "V1.0", b.strip()
    return v or "V1.0", ""


def _focus(text: str) -> str:
    return "查阅源文件、系统记录、日期/时间窗、研究者判断依据及EDC录入，确认与方案要求一致并保留可溯源证据。"


def _remove_leading_label(text: str, labels: List[str]) -> str:
    t = _strip(text)
    for lab in labels:
        t = re.sub(rf"^{re.escape(lab)}\s*[:：]?\s*", "", t)
    return t.strip()


def _split_items(text: Any, labels: List[str] | None = None) -> List[str]:
    """Split AI generated merged bullets into atomic rows for Word table cells."""
    t = _clean(text)
    if labels:
        t = _remove_leading_label(t, labels)
    t = re.sub(r"[；;]\s*(?=[•·▪■\-]|\d+[\.、）)]|[（(]?[一二三四五六七八九十]+[）)])", "\n", t)
    raw_lines: List[str] = []
    for line in t.splitlines():
        line = line.strip()
        if not line:
            continue
        # Markdown bullets and numbered bullets are row separators. Preserve inner punctuation such as mg, D-1.
        parts = re.split(r"(?<![A-Za-z0-9])(?:[•·▪■]|\d+[\.、）)]|[（(]?[一二三四五六七八九十]+[）)])\s*", line)
        if len(parts) > 1:
            raw_lines.extend(parts)
        else:
            raw_lines.append(line)
    out: List[str] = []
    for item in raw_lines:
        item = _strip(item)
        if labels:
            item = _remove_leading_label(item, labels)
        item = item.strip("；;。 ")
        if _ok(item) and item not in out:
            out.append(item)
    return out


def _pair_atomic(left: str, right: str, left_labels: List[str], left_prefix: str = "") -> List[Tuple[str, str]]:
    left_items = _split_items(left, left_labels) or [_remove_leading_label(left, left_labels)]
    right_items = _split_items(right) or ([_strip(right)] if _ok(right) else [])
    pairs: List[Tuple[str, str]] = []
    for idx, item in enumerate(left_items):
        if not _ok(item):
            continue
        focus = ""
        if len(right_items) == len(left_items):
            focus = right_items[idx]
        elif len(right_items) > 1 and idx < len(right_items):
            focus = right_items[idx]
        elif right_items:
            focus = "\n".join(right_items)
        value = f"{left_prefix}{item}" if left_prefix else item
        pairs.append((value, focus or _focus(value)))
    return pairs


def _criteria(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]], List[Dict[str, str]]]:
    criteria: List[Dict[str, str]] = []
    exclusions: List[Dict[str, str]] = []
    process: List[Dict[str, str]] = []
    for r in _rows(md, ["方案", "重点关注"]):
        left = _strip(r[0]); right = _strip(r[1]) if len(r) > 1 else ""
        if not _ok(left) or left in {"方案", "方案描述"}:
            continue
        if any(k in left for k in PROCESS_KEYS):
            for req, focus in _pair_atomic(left, right, ["随机化程序", "随机程序", "方案描述"]):
                process.append({"requirement": req, "focus": focus})
        elif left.startswith("排除标准") or "排除标准" in left[:12]:
            for criterion, focus in _pair_atomic(left, right, ["排除标准"], "排除标准："):
                exclusions.append({"criterion": criterion, "ai_focus": focus})
        elif left.startswith("入选标准") or "入选标准" in left[:12]:
            for criterion, focus in _pair_atomic(left, right, ["入选标准"], "入选标准："):
                criteria.append({"criterion": criterion, "ai_focus": focus})
    if not process:
        for r in _rows(md, ["方案描述", "重点关注"]):
            left = _strip(r[0]); right = _strip(r[1]) if len(r) > 1 else ""
            if _ok(left) and left != "方案描述":
                for req, focus in _pair_atomic(left, right, ["方案描述"]):
                    process.append({"requirement": req, "focus": focus})
    return criteria, exclusions, process


def _endpoint_pairs(objective: str, endpoint: str) -> List[Dict[str, str]]:
    objectives = _split_items(objective, ["主要目的", "次要目的", "探索性目的"]) or [_strip(objective)]
    endpoints = _split_items(endpoint, ["主要终点", "次要终点", "探索性终点"]) or [_strip(endpoint)]
    rows: List[Dict[str, str]] = []
    if len(objectives) == len(endpoints):
        for obj, end in zip(objectives, endpoints):
            if _ok(obj) and _ok(end):
                rows.append({"objective": obj, "endpoint": end})
        return rows
    for obj in objectives:
        if not _ok(obj):
            continue
        for end in endpoints:
            if _ok(end):
                rows.append({"objective": obj, "endpoint": end})
    return rows


def _endpoints(md: str) -> Tuple[List[Dict[str, str]], List[Dict[str, str]]]:
    primary: List[Dict[str, str]] = []
    secondary: List[Dict[str, str]] = []
    pending = ""
    for r in _rows(md, ["目的", "终点"]):
        obj = _strip(r[0]) if r else ""; end = _strip(r[1]) if len(r) > 1 else ""
        if end in {"主要终点", "次要终点", "探索性终点"}:
            if "次要" in end and _ok(obj):
                pending = obj
            continue
        if not _ok(obj) and pending and _ok(end):
            secondary.extend(_endpoint_pairs(pending, end)); pending = ""; continue
        if not (_ok(obj) and _ok(end)):
            continue
        target = secondary if any(k in obj or k in end for k in ["次要", "探索性"]) else primary
        target.extend(_endpoint_pairs(obj, end))
    return primary, secondary


def _summary(md: str) -> str:
    sec = _section(md, ["摘要总结", "摘要"], ["五、", "## 五", "一、项目简介"])
    return "\n".join([_strip(x) for x in sec.splitlines() if _ok(x)])


def _law_rows_from_line(line: str) -> Dict[str, str] | None:
    parts = [_strip(x) for x in re.split(r"\s*[｜|]\s*", line) if _ok(x)]
    if len(parts) < 3:
        return None
    return {
        "regulation": parts[0],
        "article": parts[1] if len(parts) > 1 else "",
        "original_text": parts[2] if len(parts) > 2 else "",
        "topic": parts[3] if len(parts) > 3 else "",
        "applicability": parts[4] if len(parts) > 4 else "",
    }


def _law_rows(md: str) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    table = _rows(md, ["法规", "条款"])
    for r in table:
        if not r or not _ok(r[0]):
            continue
        rows.append({
            "regulation": _strip(r[0]),
            "article": _strip(r[1]) if len(r) > 1 else "",
            "original_text": _strip(r[2]) if len(r) > 2 else "",
            "topic": _strip(r[3]) if len(r) > 3 else "",
            "applicability": _strip(r[4]) if len(r) > 4 else "",
        })
    if rows:
        return rows
    sec = _section(md, ["法规依据补充说明"], ["三、", "3.1", "质控流程"])
    for line in [_strip(x) for x in sec.splitlines() if _ok(x)]:
        parsed = _law_rows_from_line(line)
        if parsed:
            rows.append(parsed)
    return rows


def _law(md: str) -> str:
    rows = _law_rows(md)
    if rows:
        return "\n".join("｜".join([x for x in [r.get("regulation", ""), r.get("article", ""), r.get("original_text", ""), r.get("topic", ""), r.get("applicability", "")] if _ok(x)]) for r in rows)
    return "本项目质控应结合现行GCP、临床试验方案、研究者手册、伦理批准文件及相关指导原则进行综合判断，确保受试者权益、安全保护及数据真实、准确、完整、可溯源。"


def _sampling(md: str) -> str:
    sec = _section(md, ["中心稽查风险评估病历抽取原则", "中心质控风险评估病历抽取原则", "抽取原则"], ["中心筛选情况", "2.5.1", "2.5.2", "2.5.3", "受试者筛选", "三、"])
    return "\n".join([_strip(x) for x in sec.splitlines() if _ok(x) and "抽取原则" not in _strip(x)])


def parse_markdown_to_template_data(md_text: str) -> Dict[str, Any]:
    md = md_text.replace("\r\n", "\n")
    author_approver = _label(md, ["撰写人/审批人", "撰写人"])
    author, approver = (author_approver.split("/", 1) + ["待审批"])[:2] if "/" in author_approver else (author_approver or "苗田", "待审批")
    version_no, version_date = _version(_label(md, ["版本号/版本日期", "版本号", "版本日期"]))
    criteria, exclusions, process = _criteria(md)
    primary, secondary = _endpoints(md)
    title = _title(md)
    sponsor = _label(md, ["申办方公司名称", "申办方名称", "申办方", "申办者"])
    audit_type = _first([r"质控类型[:：]\s*([^\n]+)", r"1\.2\s*质控类型[:：]?\s*\n([^#\n]+)"], md, "中心常规质控")
    sampling = _sampling(md)
    law_rows = _law_rows(md)
    return {
        "AUTHOR": author.strip(),
        "APPROVER": approver.strip(),
        "AUDIT_COMPANY": _label(md, ["质控公司", "稽查公司"]) or "北京万宁睿和医药科技有限公司",
        "PROJECT_TITLE": title,
        "PROJECT_CODE": _first([r"方案编号[:：]\s*([A-Za-z0-9\-]+)"], md),
        "SPONSOR_NAME": sponsor,
        "VERSION_NO": version_no,
        "VERSION_DATE": version_date,
        "AUDIT_TYPE": audit_type,
        "SUMMARY_TEXT": _summary(md),
        "IMP_SPEC": _first([r"试验药物规格[^\n]*[:：]\s*([^\n]+)"], md),
        "IMP_DESCRIPTION": "",
        "SAMPLING_PRINCIPLE": sampling,
        "RISK_SAMPLING_RULE": sampling,
        "LAW_SUPPLEMENT": _law(md),
        "LAW_SUPPLEMENT_ROWS": law_rows,
        "project": {"name": title, "sponsor": sponsor, "protocol_code": "", "version_no": version_no, "version_date": version_date, "audit_type": audit_type},
        "criteria_ai_rows": criteria,
        "exclusion_ai_rows": exclusions,
        "process_requirement_rows": process,
        "primary_endpoint_rows": primary,
        "secondary_endpoint_rows": secondary,
        "law_supplement_rows": law_rows,
        "risk_analysis_rows": [],
        "subject_rows": [],
        "assignment_rows": [],
        "report_send_rows": [{"name": "待填写", "email": "待填写", "title_company": "待填写"}],
        "defect_rows": [],
    }
