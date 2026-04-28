from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import shutil
import subprocess
from typing import Any, Dict, Iterable, List, Optional

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def enrich_template_context(data: Dict[str, Any]) -> Dict[str, Any]:
    data = dict(data)
    project = dict(data.get("project", {}))
    title = project.get("name", data.get("PROJECT_TITLE", ""))

    data["PROJECT_TITLE"] = title
    data["PROJECT_TITLE_LINE1"] = title[:28]
    data["PROJECT_TITLE_LINE2"] = title[28:] if len(title) > 28 else ""
    data["SPONSOR_NAME"] = project.get("sponsor", data.get("SPONSOR_NAME", ""))
    data["PROJECT_CODE"] = project.get("protocol_code", data.get("PROJECT_CODE", ""))
    data["VERSION_NO"] = project.get("version_no", data.get("VERSION_NO", "V1.0"))
    data["VERSION_DATE"] = project.get("version_date", data.get("VERSION_DATE", ""))
    data["AUDIT_TYPE"] = project.get("audit_type", data.get("AUDIT_TYPE", "中心常规质控"))
    data.setdefault("AUTHOR", "苗田")
    data.setdefault("APPROVER", "张艳")
    data.setdefault("AUDIT_COMPANY", "北京万宁睿和医药科技有限公司")

    pa = data.get("protocol_analysis", {}) or {}
    data["PRIMARY_OBJECTIVE"] = pa.get("primary_endpoint", "")
    data["IMP_DESCRIPTION"] = data.get("IMP_DESCRIPTION", "")
    return data


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("[填写]", "待填写").strip()


def _set_cell_text_preserve_style(cell, text: Any) -> None:
    text = _safe_text(text)
    if cell.paragraphs:
        p = cell.paragraphs[0]
        if p.runs:
            p.runs[0].text = text
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(text)
        for extra_p in cell.paragraphs[1:]:
            for r in extra_p.runs:
                r.text = ""
            if not extra_p.runs:
                extra_p.text = ""
    else:
        cell.text = text


def _fill_row(row, values: Iterable[Any]) -> None:
    values = list(values)
    for i, value in enumerate(values):
        if i < len(row.cells):
            _set_cell_text_preserve_style(row.cells[i], value)


def _clone_row(table: Table, template_row_idx: int):
    new_tr = deepcopy(table.rows[template_row_idx]._tr)
    table._tbl.append(new_tr)
    return table.rows[-1]


def _trim_table_to_template_row(table: Table, template_idx: int = 1) -> None:
    while len(table.rows) > template_idx + 1:
        table._tbl.remove(table.rows[-1]._tr)


def _replace_table_rows(table: Optional[Table], rows: List[Dict[str, Any]], row_builder) -> None:
    if table is None or not rows:
        return
    template_idx = 1 if len(table.rows) > 1 else 0
    _trim_table_to_template_row(table, template_idx)
    _fill_row(table.rows[template_idx], row_builder(rows[0]))
    for item in rows[1:]:
        row = _clone_row(table, template_idx)
        _fill_row(row, row_builder(item))


def _header_text(table: Table) -> str:
    if not table.rows:
        return ""
    return "|".join(c.text.strip().replace("\n", "") for c in table.rows[0].cells)


def _find_table(doc: Document, headers: List[str], occurrence: int = 1) -> Optional[Table]:
    count = 0
    for table in doc.tables:
        h = _header_text(table)
        if all(x in h for x in headers):
            count += 1
            if count == occurrence:
                return table
    return None


def _all_paragraphs(doc: Document):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def _set_paragraph_text_preserve_style(paragraph: Paragraph, text: Any) -> None:
    text = _safe_text(text)
    if paragraph.runs:
        paragraph.runs[0].text = text
        for r in paragraph.runs[1:]:
            r.text = ""
    else:
        paragraph.add_run(text)


def _replace_labeled_paragraph(doc: Document, label: str, value: Any) -> None:
    value = _safe_text(value)
    for p in doc.paragraphs:
        if label in p.text:
            _set_paragraph_text_preserve_style(p, f"{label}{value}")
            return


def _replace_cover_title(doc, title):
    for i, p in enumerate(doc.paragraphs):
        if "中心质控计划" in p.text and i > 0:
            target = doc.paragraphs[i - 1]

            for run in target.runs:
                run.text = ""

            target.runs[0].text = title
            return


def _fill_basic_fields(doc: Document, data: Dict[str, Any]) -> None:
    title = data.get("PROJECT_TITLE", "")
    sponsor = data.get("SPONSOR_NAME", "")
    company = data.get("AUDIT_COMPANY", "北京万宁睿和医药科技有限公司")
    version = data.get("VERSION_NO", "V1.0")
    vdate = data.get("VERSION_DATE", "")
    author = data.get("AUTHOR", "")
    approver = data.get("APPROVER", "")
    audit_type = data.get("AUDIT_TYPE", "中心常规质控")

    _replace_cover_title(doc, title)

    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.replace("\n", "").replace(" ", "")
            if "申办方" in label or "申办者" in label:
                _set_cell_text_preserve_style(row.cells[1], sponsor)
            elif "质控公司" in label or "稽查公司" in label:
                _set_cell_text_preserve_style(row.cells[1], company)
            elif "版本号/版本日期" in label:
                _set_cell_text_preserve_style(row.cells[1], f"{version}/{vdate}" if vdate else version)
            elif "撰写人/审批人" in label:
                _set_cell_text_preserve_style(row.cells[1], f"{author}/{approver}" if approver else author)

    _replace_labeled_paragraph(doc, "1.1项目名称：", title)
    _replace_labeled_paragraph(doc, "1.2质控类型：", audit_type)
    _replace_labeled_paragraph(doc, "1.3申办方：", sponsor)


def _replace_placeholders(doc: Document, data: Dict[str, Any]) -> None:
    flat = {}
    for k, v in data.items():
        if not isinstance(v, (dict, list)):
            flat[k] = v
    project = data.get("project", {}) or {}
    for k, v in project.items():
        flat[f"project.{k}"] = v
    for p in _all_paragraphs(doc):
        original = p.text
        text = original
        for k, v in flat.items():
            text = text.replace("{{" + str(k) + "}}", _safe_text(v))
        if text != original:
            _set_paragraph_text_preserve_style(p, text)


def _render_markdown_data_to_template(doc: Document, data: Dict[str, Any]) -> None:
    criteria_rows = list(data.get("criteria_ai_rows", []) or []) + list(data.get("exclusion_ai_rows", []) or [])
    process_rows = list(data.get("process_requirement_rows", []) or [])
    primary_rows = list(data.get("primary_endpoint_rows", []) or [])
    secondary_rows = list(data.get("secondary_endpoint_rows", []) or [])
    endpoint_rows = primary_rows + secondary_rows

    # 1. 数据风险表
    _replace_table_rows(
        _find_table(doc, ["数据风险因素", "详细信息"]),
        data.get("risk_analysis_rows", []) or [],
        lambda r: [r.get("risk_factor", ""), r.get("detail", "")],
    )

    # 2. 受试者基本情况表
    _replace_table_rows(
        _find_table(doc, ["筛选号", "基本情况"]),
        data.get("subject_rows", []) or [],
        lambda r: [r.get("subject_id", ""), r.get("summary", r.get("protocol_version", ""))],
    )

    # 3. 质控分工表
    _replace_table_rows(
        _find_table(doc, ["序号", "质控流程", "负责质控人员"]),
        data.get("assignment_rows", []) or [],
        lambda r: [r.get("seq", ""), r.get("process", ""), r.get("assignee", ""), r.get("plan_time", "")],
    )

    # 4. 入排及方案执行表：方案 / 重点关注
    _replace_table_rows(
        _find_table(doc, ["方案", "重点关注"], occurrence=1),
        criteria_rows,
        lambda r: [r.get("criterion", ""), r.get("ai_focus", "")],
    )

    # 5. 随机化/复筛/终止等流程表：方案描述 / 重点关注
    _replace_table_rows(
        _find_table(doc, ["方案描述", "重点关注"], occurrence=1),
        process_rows,
        lambda r: [r.get("requirement", ""), r.get("focus", "")],
    )

    # 6. 研究目的和终点表
    endpoint_table = _find_table(doc, ["主要目的", "主要终点"]) or _find_table(doc, ["主要目的/次要目的", "主要终点/次要终点"])
    _replace_table_rows(
        endpoint_table,
        endpoint_rows,
        lambda r: [r.get("objective", r.get("purpose", "")), r.get("endpoint", "")],
    )

    # 7. 安全性表。注意模板还有其他“类别/重点关注”表，取第一个类别+重点关注表。
    _replace_table_rows(
        _find_table(doc, ["类别", "重点关注"], occurrence=1),
        data.get("safety_focus_rows", []) or [],
        lambda r: [r.get("category", ""), r.get("focus", "")],
    )

    # 8. 试验用药品规格表
    imp_table = _find_table(doc, ["试验药物规格", "剂型"])
    imp_text = data.get("IMP_DESCRIPTION", "")
    if imp_table is not None and imp_text:
        template_idx = 1 if len(imp_table.rows) > 1 else 0
        _trim_table_to_template_row(imp_table, template_idx)
        _fill_row(imp_table.rows[template_idx], [imp_text, ""])

    # 9. 报告发送表
    _replace_table_rows(
        _find_table(doc, ["姓名", "邮箱", "职位/公司"]),
        data.get("report_send_rows", []) or [],
        lambda r: [r.get("name", ""), r.get("email", ""), r.get("title_company", "")],
    )


def generate_docx_from_template(template_path: str | Path, data: Dict[str, Any], output_path: str | Path) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)
    if not template_path.exists():
        raise FileNotFoundError(f"模板不存在：{template_path}")

    data = enrich_template_context(data)
    doc = Document(str(template_path))

    _replace_placeholders(doc, data)
    _fill_basic_fields(doc, data)
    _render_markdown_data_to_template(doc, data)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def convert_docx_to_pdf(docx_path: str | Path, pdf_path: str | Path | None = None, output_dir: str | Path | None = None) -> Path | None:
    docx_path = Path(docx_path)
    if pdf_path is not None:
        pdf_path = Path(pdf_path)
    elif output_dir is not None:
        pdf_path = Path(output_dir) / docx_path.with_suffix(".pdf").name
    else:
        pdf_path = docx_path.with_suffix(".pdf")

    soffice = shutil.which("soffice")
    if not soffice:
        return None
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)], check=True, capture_output=True, text=True)
    except Exception:
        return None

    generated = pdf_path.parent / docx_path.with_suffix(".pdf").name
    if generated.exists() and generated != pdf_path:
        try:
            generated.replace(pdf_path)
        except Exception:
            return generated
    return pdf_path if pdf_path.exists() else generated if generated.exists() else None
