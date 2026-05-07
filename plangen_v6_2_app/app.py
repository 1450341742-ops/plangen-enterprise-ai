import re
from pathlib import Path
import streamlit as st
from docx import Document

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context
from planner.core.dingtalk_agent_client import call_dingtalk_agent

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"

st.set_page_config(page_title="PlanGen 稽查计划生成", layout="centered")
st.title("PlanGen｜稽查计划一键生成")
st.caption("支持钉钉智能体自动思考生成、粘贴 Markdown、上传钉钉AI质控结果 Word，并自动映射到内置模板。")

if not TEMPLATE_PATH.exists():
    st.error("系统内置模板缺失：请将 AI_Center_QC_Plan_Template_20260310.docx 放入 plangen_v6_2_app/templates/ 目录。")
    st.stop()

input_mode = st.radio("选择输入方式", ["钉钉智能体生成", "粘贴 Markdown", "上传 Word"], horizontal=True, key="input_mode")

source_text = ""
project_name_override = ""

if input_mode == "钉钉智能体生成":
    st.subheader("钉钉智能体入口")
    st.caption("输入需求或粘贴方案摘要，系统先调用钉钉智能体生成标准Markdown，再自动映射模板。")
    project_name_override = st.text_input("项目名称（可选，用于下载文件命名）", key="agent_project_name")
    user_prompt = st.text_area("给钉钉智能体的生成要求", value="请根据以下方案内容生成中心质控计划所需的标准Markdown内容。", height=130, key="agent_prompt")
    protocol_text = st.text_area("方案内容/补充资料（可粘贴Protocol摘要、方案关键信息或钉钉知识库整理内容）", height=260, key="agent_protocol_text")
    uploaded_agent_docx = st.file_uploader("上传方案或钉钉AI结果 Word（可选）", type=["docx"], key="agent_docx_input")

    if uploaded_agent_docx:
        temp_docx = OUTPUT_DIR / uploaded_agent_docx.name
        temp_docx.write_bytes(uploaded_agent_docx.getvalue())
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
        protocol_text = protocol_text + "\n" + "\n".join(parts)
        st.success("Word内容读取完成，将作为智能体输入。")

    if st.button("调用钉钉智能体并生成稽查计划", type="primary", key="agent_generate_btn"):
        try:
            with st.spinner("钉钉智能体正在分析并生成内容..."):
                source_text = call_dingtalk_agent(user_prompt, protocol_text=protocol_text)
            st.session_state["agent_markdown_result"] = source_text
            st.success("钉钉智能体内容生成完成，正在映射模板。")
        except Exception as e:
            st.error(f"钉钉智能体调用失败：{e}")

    if st.session_state.get("agent_markdown_result"):
        source_text = st.session_state["agent_markdown_result"]
        with st.expander("查看钉钉智能体生成的Markdown", expanded=False):
            st.text_area("智能体生成结果", source_text, height=320, key="agent_result_preview")

elif input_mode == "粘贴 Markdown":
    source_text = st.text_area("粘贴钉钉生成的 Markdown 全文", height=420, key="md_input")

else:
    uploaded_docx = st.file_uploader("上传钉钉AI质控结果 Word 文档（.docx）", type=["docx"], key="docx_input")
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
        st.success("Word内容读取完成，可以生成稽查计划。")

if source_text.strip():
    st.caption(f"已读取内容长度：{len(source_text)} 字符")
    generate_label = "生成并下载稽查计划" if input_mode != "钉钉智能体生成" else "使用智能体结果生成并下载稽查计划"
    if st.button(generate_label, type="primary", key="generate_btn"):
        try:
            data = parse_markdown_to_template_data(source_text)
            data = enrich_template_context(data)
            if project_name_override:
                data["PROJECT_TITLE"] = project_name_override
                data.setdefault("project", {})["name"] = project_name_override
            project_name = data.get("project", {}).get("name", "稽查计划") or data.get("PROJECT_TITLE", "稽查计划") or "稽查计划"
            safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", project_name).strip()[:60]
            output_path = OUTPUT_DIR / f"{safe_name} 稽查计划.docx"
            generate_docx_from_template(TEMPLATE_PATH, data, output_path)
            st.success("稽查计划生成成功，请点击下方按钮下载。")
            st.download_button(
                "下载稽查计划 Word",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_btn"
            )
        except Exception as e:
            st.error(f"生成失败：{e}")
else:
    st.info("请选择钉钉智能体生成、粘贴 Markdown，或上传钉钉AI质控结果 Word 文档。")
