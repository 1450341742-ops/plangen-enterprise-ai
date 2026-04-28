import json
import os
from pathlib import Path
from datetime import date
import streamlit as st

from planner.core.parser_v62_ai import extract_text_from_file, build_v62_plan_json
from planner.core.ai_engine import check_ollama_status
from planner.core.renderer import generate_docx_from_template, convert_docx_to_pdf, enrich_template_context

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

USER_TEMPLATE = APP_DIR / "templates" / "AI_Center_QC_Plan_Template_20260310.docx"
DEFAULT_TEMPLATE = USER_TEMPLATE if USER_TEMPLATE.exists() else APP_DIR / "templates" / "PlanGen_v62_Template.docx"

IS_CLOUD = bool(os.getenv("STREAMLIT_SERVER_PORT") or os.getenv("STREAMLIT_SHARING_MODE") or os.getenv("HOSTNAME"))

st.set_page_config(page_title="PlanGen 团队内部版", layout="wide")
st.title("PlanGen 团队内部版｜OCR + AI/规则双引擎")
st.caption("支持扫描PDF OCR、可复制PDF、Word和TXT。团队本地部署可使用OCR和Ollama；云端会自动降级为规则引擎。")

with st.sidebar:
    st.subheader("运行模式")
    if IS_CLOUD:
        st.warning("当前为云端环境：本地Ollama不可用；OCR取决于云端是否安装Tesseract。建议团队内部用本地部署版。")
        mode = "仅规则兜底"
    else:
        mode = st.radio("选择模式", ["规则+AI增强（本地Ollama）", "仅规则兜底"], index=1)

    st.markdown(
        "**团队内部版能力：**\n"
        "- 扫描PDF OCR\n"
        "- 可复制PDF/Word/TXT解析\n"
        "- 严格项目字段提取\n"
        "- 入排标准章节切块\n"
        "- 行业级重点关注生成\n"
        "- JSON预览与编辑\n"
        "- 自动写入你的Word模板\n"
        "- DOCX导出"
    )

    if not IS_CLOUD:
        st.markdown(
            "**本地OCR/AI准备：**\n"
            "```powershell\n"
            "tesseract --version\n"
            "ollama pull qwen2.5:3b\n"
            "ollama serve\n"
            "```"
        )

ollama_ok, ollama_msg = (False, "当前未启用本地AI")
if not IS_CLOUD and mode == "规则+AI增强（本地Ollama）":
    ollama_ok, ollama_msg = check_ollama_status()
    if ollama_ok:
        st.success(ollama_msg)
    else:
        st.warning(f"{ollama_msg}。系统会自动降级为规则引擎。")

uploaded_protocol = st.file_uploader("上传方案（扫描PDF / PDF / DOCX / TXT）", type=["pdf", "docx", "txt"])
manual_text = st.text_area("可选：粘贴OCR后的方案文本（如PDF识别效果差，可直接粘贴WPS/Adobe OCR文本）", height=180)
uploaded_template = st.file_uploader("上传Word模板（可选；不上传则使用系统内置的你的中心质控计划模板）", type=["docx"])

template_path = DEFAULT_TEMPLATE
if uploaded_template:
    template_path = OUTPUT_DIR / uploaded_template.name
    template_path.write_bytes(uploaded_template.getvalue())

if USER_TEMPLATE.exists() and template_path == USER_TEMPLATE:
    st.info("当前模板：已使用你的中心质控计划正式模板 AI_Center_QC_Plan_Template_20260310.docx")
else:
    st.info(f"当前模板：{template_path.name}")

raw_text = ""
if uploaded_protocol:
    raw_text = extract_text_from_file(uploaded_protocol)

if manual_text and len(manual_text.strip()) > 50:
    raw_text = manual_text

if uploaded_protocol or manual_text:
    st.subheader("协议文本预览")
    st.text_area("提取/粘贴结果", raw_text[:30000], height=300)

    if raw_text.startswith("文件解析失败") or raw_text.startswith("PDF解析失败"):
        st.error(raw_text)
    elif len(raw_text.strip()) < 300:
        st.error("提取到的文本过少，无法生成有效质控计划。请确认：1）本地已安装Tesseract；2）扫描PDF清晰；3）或将WPS/Adobe OCR后的文本粘贴到上方文本框。")
    else:
        if st.button("生成结构化JSON"):
            use_ai = (not IS_CLOUD) and mode == "规则+AI增强（本地Ollama）" and ollama_ok
            if mode == "规则+AI增强（本地Ollama）" and not ollama_ok:
                st.warning("本地AI不可用，系统已自动切换为规则引擎生成。")
            with st.spinner("系统正在解析协议并生成结构化JSON，请稍候..."):
                data = build_v62_plan_json(raw_text, use_ai=use_ai)
            if isinstance(data, dict) and data.get("_error"):
                st.error(data["_error"])
                if data.get("_detail"):
                    st.code(str(data["_detail"]))
            else:
                st.session_state["plan_json"] = data
                st.success("JSON生成成功")

if "plan_json" in st.session_state:
    data = st.session_state["plan_json"]
    st.subheader("字段确认页（生成前请先确认）")

    c1, c2, c3 = st.columns(3)
    with c1:
        project_name = st.text_input("项目名称", value=data.get("project", {}).get("name", ""))
        sponsor_name = st.text_input("申办方", value=data.get("project", {}).get("sponsor", ""))
        protocol_code = st.text_input("方案编号", value=data.get("project", {}).get("protocol_code", ""))
    with c2:
        version_no = st.text_input("版本号", value=data.get("project", {}).get("version_no", "V1.0"))
        version_date = st.text_input("版本日期", value=data.get("project", {}).get("version_date", date.today().strftime("%Y年%m月%d日")))
        audit_type = st.text_input("质控类型", value=data.get("project", {}).get("audit_type", "中心常规质控"))
    with c3:
        author = st.text_input("撰写人", value=data.get("AUTHOR", "苗田"))
        approver = st.text_input("审批人", value=data.get("APPROVER", "张艳"))
        audit_company = st.text_input("质控公司", value=data.get("AUDIT_COMPANY", "北京万宁睿和医药科技有限公司"))

    data.setdefault("project", {})
    data["project"]["name"] = project_name
    data["project"]["sponsor"] = sponsor_name
    data["project"]["protocol_code"] = protocol_code
    data["project"]["version_no"] = version_no
    data["project"]["version_date"] = version_date
    data["project"]["audit_type"] = audit_type
    data["AUTHOR"] = author
    data["APPROVER"] = approver
    data["AUDIT_COMPANY"] = audit_company

    edited_text = st.text_area("JSON预览与编辑", value=json.dumps(data, ensure_ascii=False, indent=2), height=560)

    tabs = st.tabs(["项目概览", "入排标准", "流程要求", "重点关注", "RBQM", "访谈问题", "发现/CAPA草案", "完整JSON"])
    parsed = data

    with tabs[0]:
        st.write(parsed.get("project", {}))
        st.write(parsed.get("protocol_analysis", {}))
    with tabs[1]:
        st.write("入选标准")
        st.dataframe(parsed.get("criteria_ai_rows", []), use_container_width=True)
        st.write("排除标准")
        st.dataframe(parsed.get("exclusion_ai_rows", []), use_container_width=True)
    with tabs[2]:
        st.dataframe(parsed.get("process_requirement_rows", []), use_container_width=True)
    with tabs[3]:
        st.text(parsed.get("ai_audit_key_points", ""))
    with tabs[4]:
        st.text(parsed.get("rbqm_strategy", ""))
    with tabs[5]:
        st.text(parsed.get("interview_questions", ""))
    with tabs[6]:
        st.text(parsed.get("finding_capa_draft", ""))
    with tabs[7]:
        st.code(json.dumps(parsed, ensure_ascii=False, indent=2), language="json")

    if st.button("生成正式文档"):
        try:
            edited_data = json.loads(edited_text)
            edited_data = enrich_template_context(edited_data)
            docx_path = OUTPUT_DIR / "PlanGen_Internal_OCR_Output.docx"
            generate_docx_from_template(template_path=template_path, data=edited_data, output_path=docx_path)
            st.success("DOCX 生成成功")
            st.download_button("下载 DOCX", data=docx_path.read_bytes(), file_name=docx_path.name)
            pdf_path = convert_docx_to_pdf(docx_path, output_dir=OUTPUT_DIR)
            if pdf_path and Path(pdf_path).exists():
                st.download_button("下载 PDF", data=Path(pdf_path).read_bytes(), file_name=Path(pdf_path).name)
            else:
                st.info("如需PDF，可下载DOCX后本地另存为PDF。")
        except Exception as e:
            st.error(f"正式生成文档失败：{e}")
