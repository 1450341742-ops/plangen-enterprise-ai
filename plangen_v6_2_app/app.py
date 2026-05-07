import re
from pathlib import Path
import streamlit as st

from planner.core.markdown_mapper import parse_markdown_to_template_data
from planner.core.renderer import generate_docx_from_template, enrich_template_context
from planner.core.llm_client import call_llm
from planner.core.knowledge_ingestor import save_uploaded_knowledge, list_knowledge_documents
from planner.core.vector_store import build_vector_index, vector_index_status
from planner.core.docx_reader import uploaded_file_to_text

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

TEMPLATE_PATH = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"

st.set_page_config(page_title="PlanGen 稽查计划生成", layout="wide")
st.title("PlanGen｜V9.0 AI稽查总监平台")
st.caption("支持 DeepSeek / Qwen / 硅基流动、企业RAG、长期记忆、PDF/Word方案解析、自动模板映射。")

with st.sidebar:
    st.header("企业知识库")

    uploaded_kb = st.file_uploader(
        "上传知识库文件",
        type=["md", "txt", "docx", "pdf"],
        key="kb_upload"
    )

    if uploaded_kb:
        try:
            saved = save_uploaded_knowledge(uploaded_kb.name, uploaded_kb.getvalue())
            st.success(f"知识文件已导入：{saved.name}")
        except Exception as e:
            st.error(f"知识导入失败：{e}")

    if st.button("重建向量知识库"):
        with st.spinner("正在构建企业向量知识库..."):
            build_vector_index(force=True)
        st.success("企业向量知识库重建完成")

    status = vector_index_status()
    st.caption(f"知识块数量：{status.get('chunk_count', 0)}")

    docs = list_knowledge_documents()
    if docs:
        st.markdown("### 已学习知识文件")
        for d in docs:
            st.markdown(f"- {d}")

if not TEMPLATE_PATH.exists():
    st.error("系统内置模板缺失")
    st.stop()

input_mode = st.radio(
    "选择输入方式",
    ["AI智能生成", "粘贴 Markdown", "上传方案文件"],
    horizontal=True,
)

source_text = ""
project_name_override = ""

if input_mode == "AI智能生成":
    st.subheader("AI智能生成入口")

    project_name_override = st.text_input("项目名称")

    user_prompt = st.text_area(
        "AI生成要求",
        value="请根据以下方案内容生成中心质控计划所需的标准Markdown内容。",
        height=120,
    )

    protocol_text = st.text_area(
        "方案内容/Protocol摘要",
        height=260,
    )

    uploaded_protocol = st.file_uploader(
        "上传方案文件（支持PDF/DOCX/TXT/MD）",
        type=["docx", "pdf", "txt", "md"],
    )

    if uploaded_protocol:
        try:
            temp_path = OUTPUT_DIR / uploaded_protocol.name
            temp_path.write_bytes(uploaded_protocol.getvalue())
            extracted_text = uploaded_file_to_text(temp_path)
            protocol_text = protocol_text + "\n\n" + extracted_text
            st.success(f"文件解析完成：{uploaded_protocol.name}")
        except Exception as e:
            st.error(f"文件解析失败：{e}")

    if st.button("调用企业AI生成稽查计划", type="primary"):
        try:
            with st.spinner("AI正在结合企业知识库进行分析..."):
                source_text = call_llm(user_prompt, protocol_text=protocol_text)
            st.session_state["agent_markdown_result"] = source_text
            st.success("企业AI内容生成完成")
        except Exception as e:
            st.error(f"AI调用失败：{e}")

    if st.session_state.get("agent_markdown_result"):
        source_text = st.session_state["agent_markdown_result"]
        with st.expander("查看AI生成Markdown"):
            st.text_area("AI生成结果", source_text, height=320)

elif input_mode == "粘贴 Markdown":
    source_text = st.text_area("粘贴AI生成Markdown", height=420)

else:
    uploaded_file = st.file_uploader(
        "上传方案文件（支持PDF/DOCX/TXT/MD）",
        type=["docx", "pdf", "txt", "md"]
    )

    if uploaded_file:
        try:
            temp_path = OUTPUT_DIR / uploaded_file.name
            temp_path.write_bytes(uploaded_file.getvalue())
            source_text = uploaded_file_to_text(temp_path)
            st.success(f"文件解析完成：{uploaded_file.name}")
        except Exception as e:
            st.error(f"文件解析失败：{e}")

if source_text.strip():
    st.caption(f"已读取内容长度：{len(source_text)} 字符")

    if st.button("生成并下载稽查计划", type="primary"):
        try:
            data = parse_markdown_to_template_data(source_text)
            data = enrich_template_context(data)

            if project_name_override:
                data["PROJECT_TITLE"] = project_name_override
                data.setdefault("project", {})["name"] = project_name_override

            project_name = data.get("project", {}).get("name", "稽查计划") or data.get("PROJECT_TITLE", "稽查计划")

            safe_name = re.sub(r"[\\/:*?\"<>|\r\n]+", "_", project_name).strip()[:60]

            output_path = OUTPUT_DIR / f"{safe_name} 稽查计划.docx"

            generate_docx_from_template(TEMPLATE_PATH, data, output_path)

            st.success("稽查计划生成成功")

            st.download_button(
                "下载稽查计划 Word",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"生成失败：{e}")
else:
    st.info("请选择AI智能生成、粘贴Markdown或上传方案文件")
