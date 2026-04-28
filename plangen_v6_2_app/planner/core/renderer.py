from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import shutil
import subprocess
from typing import Any, Dict, Iterable

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

def flatten_json(data: Dict[str, Any], parent_key: str = "") -> Dict[str, Any]:
    items: Dict[str, Any] = {}
    for k, v in data.items():
        new_key = f"{parent_key}.{k}" if parent_key else str(k)
        if isinstance(v, dict):
            items.update(flatten_json(v, new_key))
        else:
            items[new_key] = v
    return items

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
    data["AUDIT_TYPE"] = project.get("audit_type", data.get("AUDIT_TYPE", "中心常规稽查"))
    data.setdefault("AUTHOR", "苗田")
    data.setdefault("APPROVER", "张艳")
    data.setdefault("AUDIT_COMPANY", "北京万宁睿和医药科技有限公司")

    pa = data.get("protocol_analysis", {})
    data["PRIMARY_OBJECTIVE"] = pa.get("primary_endpoint", "")
    data["IMP_DESCRIPTION"] = "待补充"
    return data

def _replace_text_in_paragraph(paragraph: Paragraph, mapping: Dict[str, Any]) -> None:
    text = paragraph.text
    replaced = False
    for key, value in mapping.items():
        placeholder = "{{" + str(key) + "}}"
        if placeholder in text:
            text = text.replace(placeholder, "" if value is None else str(value))
            replaced = True
    if replaced:
        if paragraph.runs:
            paragraph.runs[0].text = text
            for run in paragraph.runs[1:]:
                run.text = ""
        else:
            paragraph.add_run(text)

def replace_placeholders(doc: Document, mapping: Dict[str, Any]) -> None:
    for p in doc.paragraphs:
        _replace_text_in_paragraph(p, mapping)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_text_in_paragraph(p, mapping)

def _table_header_text(table: Table) -> str:
    try:
        return " | ".join(c.text.strip() for c in table.rows[0].cells)
    except Exception:
        return ""

def _add_row_by_copy(table: Table, template_row_idx: int):
    tr = deepcopy(table.rows[template_row_idx]._tr)
    table._tbl.append(tr)
    return table.rows[-1]

def _fill_row(row, values: Iterable[Any]) -> None:
    values = list(values)
    for i, val in enumerate(values):
        if i >= len(row.cells):
            break
        text = "" if val is None else str(val)
        cell = row.cells[i]
        if cell.paragraphs:
            cell.paragraphs[0].text = text
            for p in cell.paragraphs[1:]:
                p.text = ""
        else:
            cell.text = text

def populate_dynamic_tables(doc: Document, data: Dict[str, Any]) -> None:
    config = [
        ("数据风险因素", "risk_analysis_rows", lambda r: [r.get("risk_factor", ""), r.get("detail", "")]),
        ("筛选号", "subject_rows", lambda r: [r.get("subject_id", ""), r.get("protocol_version", ""), r.get("summary", "")]),
        ("序号", "assignment_rows", lambda r: [r.get("seq", ""), r.get("process", ""), r.get("assignee", ""), r.get("plan_time", "")]),
        ("方案", "criteria_ai_rows", lambda r: [r.get("criterion", ""), r.get("ai_focus", "")]),
        ("方案描述", "process_requirement_rows", lambda r: [r.get("requirement", ""), r.get("focus", "")]),
        ("次要目的", "secondary_endpoint_rows", lambda r: [r.get("objective", ""), r.get("endpoint", "")]),
        ("类别", "safety_focus_rows", lambda r: [r.get("category", ""), r.get("focus", "")]),
        ("分类", "defect_rows", lambda r: [r.get("category", ""), r.get("yes", ""), r.get("no", ""), r.get("minor", ""), r.get("major", "")]),
        ("姓名", "report_send_rows", lambda r: [r.get("name", ""), r.get("email", ""), r.get("title_company", "")]),
    ]
    used_tables = set()
    for table in doc.tables:
        header = _table_header_text(table)
        for header_keyword, data_key, row_builder in config:
            if header_keyword in header and id(table) not in used_tables:
                rows = data.get(data_key, [])
                if not isinstance(rows, list) or not rows:
                    used_tables.add(id(table))
                    break
                template_idx = 1 if len(table.rows) > 1 else 0
                while len(table.rows) > template_idx + 1:
                    table._tbl.remove(table.rows[-1]._tr)
                _fill_row(table.rows[template_idx], row_builder(rows[0]))
                for item in rows[1:]:
                    new_row = _add_row_by_copy(table, template_idx)
                    _fill_row(new_row, row_builder(item))
                used_tables.add(id(table))
                break

def generate_docx_from_template(template_path: str | Path, data: Dict[str, Any], output_path: str | Path) -> Path:
    template_path = Path(template_path)
    output_path = Path(output_path)
    data = enrich_template_context(data)
    doc = Document(str(template_path))
    flat = flatten_json(data)
    replace_placeholders(doc, flat)
    replace_placeholders(doc, {k: v for k, v in data.items() if not isinstance(v, (dict, list))})
    populate_dynamic_tables(doc, data)
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
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent), str(docx_path)],
            check=True, capture_output=True, text=True
        )
    except Exception:
        return None

    generated = pdf_path.parent / docx_path.with_suffix(".pdf").name
    if generated.exists() and generated != pdf_path:
        try:
            generated.replace(pdf_path)
        except Exception:
            return generated
    if pdf_path.exists():
        return pdf_path
    if generated.exists():
        return generated
    return None
