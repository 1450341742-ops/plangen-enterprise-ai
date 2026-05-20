import re
from pathlib import Path
import streamlit as st

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.adaptive_template_engine import adaptive_map_template
from planner.core.renderer import enrich_template_context
from planner.core.summary_postprocess import extract_summary_fallback, ensure_summary_written

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = APP_DIR / "templates" / "中心质控计划_占位符模板.docx"

st.set_page_config(page_title="PlanGen 中心质控计划生成", layout="centered")

st.title("PlanGen｜中心质控计划生成系统")
st.caption("复制粘贴 Markdown 或钉钉质控计划初稿，自动映射到内置模板并下载Word。")

if not TEMPLATE_PATH.exists():
    st.error("系统内置模板缺失，请将 中心质控计划_占位符模板.docx 放入 templates 目录。")
    st.stop()

source_text = st.text_area(
    "粘贴 Markdown / 钉钉质控计划初稿",
    height=560,
    placeholder="请将钉钉AI生成的中心质控计划初稿或Markdown内容粘贴到这里。"
)

project_name_override = st.text_input("项目名称（可选，用于修正封面及下载文件名）")

if source_text.strip():
    st.caption(f"已读取内容长度：{len(source_text)} 字符")

    if st.button("生成并下载中心质控计划", type="primary"):
        try:
            data = parse_markdown_to_template_data(source_text)
            data["RAW_MARKDOWN"] = source_text
            if not str(data.get("SUMMARY_TEXT") or "").strip():
                data["SUMMARY_TEXT"] = extract_summary_fallback(source_text)
            data = enrich_template_context(data)

            if project_name_override.strip():
                data["PROJECT_TITLE"] = project_name_override.strip()
                data.setdefault("project", {})["name"] = project_name_override.strip()

            project_name = data.get("PROJECT_TITLE") or data.get("project", {}).get("name") or "中心质控计划"
            safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", project_name).strip()[:60]
            output_path = OUTPUT_DIR / f"{safe_name} 中心质控计划.docx"

            adaptive_map_template(TEMPLATE_PATH, data, output_path)
            ensure_summary_written(output_path, data.get("SUMMARY_TEXT", ""))

            st.success("中心质控计划生成成功，请点击下方按钮下载。")
            st.download_button(
                "下载中心质控计划 Word",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"生成失败：{e}")
else:
    st.info("请先粘贴 Markdown 或钉钉质控计划初稿。")