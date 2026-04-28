import re
from pathlib import Path
import streamlit as st
from docx import Document

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"

st.set_page_config(page_title="PlanGen 稽查计划生成", layout="wide")
st.title("PlanGen｜稽查计划一键生成")
st.caption("复制粘贴钉钉 Markdown，或上传钉钉AI质控结果 Word 文档，系统自动映射到内置模板并下载生成文件。")

if not TEMPLATE_PATH.exists():
    st.error("系统内置模板缺失：请将 AI_Center_QC_Plan_Template_20260310.docx 放入 plangen_v6_2_app/templates/ 目录。")
    st.stop()

input_mode = st.radio("选择输入方式", ["粘贴 Markdown", "上传钉钉AI质控结果 Word"], horizontal=True)

source_text = ""
if input_mode == "粘贴 Markdown":
    source_text = st.text_area("粘贴钉钉生成的 Markdown 全文", height=430)
else:
    uploaded_docx = st.file_uploader("上传钉钉AI质控结果 Word 文档（.docx）", type=["docx"])
    if uploaded_docx:
        temp_docx = OUTPUT_DIR / uploaded_docx.name
        temp_docx.write_bytes(uploaded_docx.getvalue())
        doc = Document(str(temp_docx))
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                if any(cells):
                    parts.append("| " + " | ".join(cells) + " |")
        source_text = "\n".join(parts)

if source_text.strip():
    with st.expander("查看输入内容", expanded=False):
        st.text_area("输入内容预览", source_text[:60000], height=260)

    if st.button("生成稽查计划", type="primary"):
        try:
            data = parse_markdown_to_template_data(source_text)
            data = enrich_template_context(data)
            project_name = data.get("project", {}).get("name", "稽查计划") or "稽查计划"
            safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", project_name).strip()[:60]
            output_path = OUTPUT_DIR / f"{safe_name} 稽查计划.docx"
            generate_docx_from_template(TEMPLATE_PATH, data, output_path)
            st.success("稽查计划生成成功")
            st.download_button("下载稽查计划 Word", data=output_path.read_bytes(), file_name=output_path.name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"生成失败：{e}")
else:
    st.info("请粘贴 Markdown 内容，或上传钉钉AI质控结果 Word 文档。")
