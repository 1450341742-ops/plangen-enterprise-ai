from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn

KEEP_MARKS = ["保持以下文字不变", "以下文字不变", "固定内容", "不需填写", "无需填写"]
PLACEHOLDER_MARKS = ["请填写", "待填写", "xxxxx", "XXXX", "{{"]

FIELD_ALIASES = {
    "项目名称": ["PROJECT_TITLE", "project.name"],
    "研究题目": ["PROJECT_TITLE", "project.name"],
    "方案标题": ["PROJECT_TITLE", "project.name"],
    "申办方": ["SPONSOR_NAME", "project.sponsor"],
    "申办者": ["SPONSOR_NAME", "project.sponsor"],
    "方案编号": ["PROJECT_CODE", "project.protocol_code"],
    "Protocol": ["PROJECT_CODE", "project.protocol_code"],
    "质控类型": ["AUDIT_TYPE", "project.audit_type"],
    "版本号/版本日期": ["VERSION_FULL"],
    "版本号": ["VERSION_NO", "project.version_no"],
    "版本日期": ["VERSION_DATE", "project.version_date"],
    "摘要总结": ["SUMMARY_TEXT"],
    "抽取原则": ["RISK_SAMPLING_RULE", "SAMPLING_PRINCIPLE"],
    "中心稽查风险评估病历抽取原则": ["RISK_SAMPLING_RULE", "SAMPLING_PRINCIPLE"],
    "试验药物规格": ["IMP_SPEC", "IMP_DESCRIPTION"],
    "药品规格": ["IMP_SPEC", "IMP_DESCRIPTION"],
    "法规依据补充说明": ["LAW_SUPPLEMENT"],
    "法规依据": ["LAW_SUPPLEMENT"],
    "质控公司": ["AUDIT_COMPANY"],
    "稽查公司": ["AUDIT_COMPANY"],
    "撰写人": ["AUTHOR"],
    "审批人": ["APPROVER"],
}

TABLE_TYPES = {
    "criteria": ["方案", "重点关注"],
    "process": ["方案描述", "重点关注"],
    "endpoint": ["主要目的", "主要终点"],
    "safety": ["类别", "重点关注"],
    "risk": ["数据风险因素", "详细信息"],
    "subject": ["筛选号", "基本情况"],
}


def _get_value(data: Dict[str, Any], key: str) -> str:
    if key == "VERSION_FULL":
        v = _get_value(data, "VERSION_NO") or _get_value(data, "project.version_no")
        d = _get_value(data, "VERSION_DATE") or _get_value(data, "project.version_date")
        return f"{v}/{d}" if v and d else v or d
    cur: Any = data
    for part in key.split("."):
        if isinstance(cur, dict):
            cur = cur.get(part)
        else:
            return ""
    return "" if cur is None else str(cur).strip()


def _best_value(label: str, data: Dict[str, Any]) -> str:
    label = (label or "").replace("\n", "").replace("：", "").replace(":", "").strip()
    for alias, keys in FIELD_ALIASES.items():
        if alias in label:
            for key in keys:
                value = _get_value(data, key)
                if value:
                    return value
    return ""


def _is_keep_text(text: str) -> bool:
    return any(x in (text or "") for x in KEEP_MARKS)


def _has_placeholder(text: str) -> bool:
    return any(x in (text or "") for x in PLACEHOLDER_MARKS)


def _set_run_style_like(run, size: Optional[Pt] = None, bold: Optional[bool] = None) -> None:
    run.font.name = run.font.name or "宋体"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold


def _set_para_text(p, text: str, keep_style: bool = True, cover_title: bool = False) -> None:
    if not p.runs:
        p.add_run("")
    if cover_title:
        p.runs[0].text = text
        _set_run_style_like(p.runs[0], Pt(18), True)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        old_size = p.runs[0].font.size
        old_bold = p.runs[0].font.bold
        p.runs[0].text = text
        if keep_style:
            _set_run_style_like(p.runs[0], old_size, old_bold)
    for r in p.runs[1:]:
        r.text = ""


def _set_cell_text(cell, text: str) -> None:
    if not cell.paragraphs:
        cell.text = text
        return
    _set_para_text(cell.paragraphs[0], text, keep_style=True)
    for p in cell.paragraphs[1:]:
        for r in p.runs:
            r.text = ""


def _table_text(table) -> str:
    return "\n".join("|".join(c.text.strip() for c in row.cells) for row in table.rows)


def _infer_table_type(table) -> str:
    text = _table_text(table)
    for name, headers in TABLE_TYPES.items():
        if all(h in text for h in headers):
            return name
    return "unknown"


def _clone_row(table, idx: int):
    table._tbl.append(deepcopy(table.rows[idx]._tr))
    return table.rows[-1]


def _rows_for_type(data: Dict[str, Any], table_type: str) -> List[Dict[str, Any]]:
    if table_type == "criteria":
        return list(data.get("criteria_ai_rows", []) or []) + list(data.get("exclusion_ai_rows", []) or [])
    if table_type == "process":
        return list(data.get("process_requirement_rows", []) or [])
    if table_type == "endpoint":
        return list(data.get("primary_endpoint_rows", []) or []) + list(data.get("secondary_endpoint_rows", []) or [])
    if table_type == "safety":
        return list(data.get("safety_focus_rows", []) or [])
    if table_type == "risk":
        return list(data.get("risk_analysis_rows", []) or [])
    if table_type == "subject":
        return list(data.get("subject_rows", []) or [])
    return []


def _values_for_type(item: Dict[str, Any], table_type: str) -> List[str]:
    if table_type == "criteria":
        return [item.get("criterion", ""), item.get("ai_focus", "")]
    if table_type == "process":
        return [item.get("requirement", ""), item.get("focus", "")]
    if table_type == "endpoint":
        return [item.get("objective", item.get("purpose", "")), item.get("endpoint", "")]
    if table_type == "safety":
        return [item.get("category", ""), item.get("focus", "")]
    if table_type == "risk":
        return [item.get("risk_factor", ""), item.get("detail", "")]
    if table_type == "subject":
        return [item.get("subject_id", ""), item.get("summary", "")]
    return []


def _fill_dynamic_table(table, rows: List[Dict[str, Any]], table_type: str) -> None:
    if not rows or not table.rows:
        return
    template_idx = 1 if len(table.rows) > 1 else 0
    while len(table.rows) > template_idx + 1:
        table._tbl.remove(table.rows[-1]._tr)
    while len(table.rows) < template_idx + 1 + len(rows):
        _clone_row(table, template_idx)
    for row_obj, item in zip(table.rows[template_idx:template_idx + len(rows)], rows):
        vals = _values_for_type(item, table_type)
        for i, val in enumerate(vals):
            if i < len(row_obj.cells):
                _set_cell_text(row_obj.cells[i], str(val))


def _find_cover_title_paragraph(doc: Document):
    for i, p in enumerate(doc.paragraphs):
        if "中心质控计划" in p.text or "稽查计划" in p.text:
            if i > 0:
                return doc.paragraphs[i - 1]
    return None


def adaptive_map_template(template_path: str | Path, data: Dict[str, Any], output_path: str | Path) -> Path:
    doc = Document(str(template_path))
    project_title = _get_value(data, "PROJECT_TITLE") or _get_value(data, "project.name")

    # 封面标题自适应
    cover_p = _find_cover_title_paragraph(doc)
    if cover_p is not None and project_title:
        _set_para_text(cover_p, project_title, cover_title=True)

    # 段落占位符自适应替换：固定内容不动
    for p in doc.paragraphs:
        text = p.text
        if not _has_placeholder(text) or _is_keep_text(text):
            continue
        if "{{" in text and "}}" in text:
            new_text = text
            for key, val in data.items():
                if isinstance(val, (str, int, float)):
                    new_text = new_text.replace("{{" + key + "}}", str(val))
            if new_text != text:
                _set_para_text(p, new_text)
                continue
        label = text.split("请填写")[0].split("待填写")[0].split("xxxxx")[0].split("XXXX")[0]
        value = _best_value(label, data)
        if value:
            new_text = text.replace("请填写", value).replace("待填写", value).replace("xxxxx", value).replace("XXXX", value)
            _set_para_text(p, new_text)

    # 表格自适应：循环表 + 字段表
    for table in doc.tables:
        if _is_keep_text(_table_text(table)):
            continue
        t_type = _infer_table_type(table)
        rows = _rows_for_type(data, t_type)
        if rows:
            _fill_dynamic_table(table, rows, t_type)
            continue
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                if not _has_placeholder(cell.text):
                    continue
                label = row.cells[idx - 1].text if idx > 0 else cell.text
                value = _best_value(label, data)
                if value:
                    _set_cell_text(cell, value)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
