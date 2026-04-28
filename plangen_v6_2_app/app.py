import json
from pathlib import Path
import streamlit as st

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

DEFAULT_TEMPLATE = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"
FALLBACK_TEMPLATE = APP_DIR / "templates" / "PlanGen_v62_Template.docx"

st.set_page_config(page_title="Markdown映射版", layout="wide")
st.title("PlanGen｜Markdown自动映射系统")
st.markdown("👉 粘贴钉钉生成的Markdown，系统自动写入模板（不做分析）")

uploaded_template = st.file_uploader("上传Word模板（必须为.docx；如仓库内已放模板可不上传）", type=["docx"])

if uploaded_template:
    template_path = OUTPUT_DIR / uploaded_template.name
    template_path.write_bytes(uploaded_template.getvalue())
    st.success(f"已使用上传模板：{uploaded_template.name}")
elif DEFAULT_TEMPLATE.exists():
    template_path = DEFAULT_TEMPLATE
    st.info(f"已使用仓库模板：{DEFAULT_TEMPLATE.name}")
elif FALLBACK_TEMPLATE.exists():
    template_path = FALLBACK_TEMPLATE
    st.warning(f"未找到正式模板，已使用备用模板：{FALLBACK_TEMPLATE.name}")
else:
    template_path = None
    st.error("未找到Word模板。请在页面上传.docx模板，或将模板放到 plangen_v6_2_app/templates/ 目录。")

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
        if template_path is None or not Path(template_path).exists():
            st.error("无法生成Word：未找到模板。请先上传.docx模板。")
        else:
            data = json.loads(edited)
            data = enrich_template_context(data)
            output = OUTPUT_DIR / "质控计划.docx"
            generate_docx_from_template(template_path, data, output)
            st.success("生成成功")
            st.download_button("下载Word", data=output.read_bytes(), file_name=output.name)
