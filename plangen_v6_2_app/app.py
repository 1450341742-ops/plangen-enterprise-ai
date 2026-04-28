import json
from pathlib import Path
import streamlit as st

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"

st.set_page_config(page_title="Markdown映射版", layout="wide")
st.title("PlanGen｜Markdown自动映射系统")

st.markdown("👉 粘贴钉钉生成的Markdown，系统自动写入模板（不做分析）")

md_text = st.text_area("粘贴Markdown全文", height=500)

if md_text:
    if st.button("解析并映射"):
        data = parse_markdown_to_template_data(md_text)
        data = enrich_template_context(data)
        st.session_state["data"] = data

if "data" in st.session_state:
    data = st.session_state["data"]

    st.subheader("结构化结果（可修改）")
    edited = st.text_area(
        "JSON编辑",
        value=json.dumps(data, ensure_ascii=False, indent=2),
        height=500
    )

    if st.button("生成Word"):
        data = json.loads(edited)

        output = OUTPUT_DIR / "质控计划.docx"
        generate_docx_from_template(TEMPLATE, data, output)

        st.success("生成成功")
        st.download_button("下载Word", data=output.read_bytes(), file_name=output.name)
