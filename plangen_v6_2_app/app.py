import json
from pathlib import Path
from datetime import date
import streamlit as st

from planner.core.parser_v62_ai import extract_text_from_file, build_v62_plan_json
from planner.core.ai_engine import check_ollama_status
from planner.core.renderer import generate_docx_from_template, convert_docx_to_pdf, enrich_template_context

APP_DIR = Path(__file__).parent
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

DEFAULT_TEMPLATE = APP_DIR / "templates" / "PlanGen_v62_Template.docx"

st.set_page_config(page_title="PlanGen v6.3 企业级版", layout="wide")
st.title("PlanGen v6.3 企业级 AI 版")
st.caption("本地免费 AI：Ollama + Qwen2.5:3b。已修复长时间卡住、AI失败中断、字段映射残留、动态表格写入不稳定等问题。")

with st.sidebar:
    st.subheader("模式")
    mode = st.radio("选择模式", ["规则+AI增强（推荐）", "仅规则兜底"], index=0)
    st.markdown(
        "**V6.3能力：**\n"
        "- 基础字段锁死提取\n"
        "- 章节切块后再AI识别\n"
        "- 入排标准 + 重点关注一次性返回\n"
        "- AI失败自动回退，不阻断文档生成\n"
        "- 生成前字段确认页\n"
        "- Word动态表格稳定写入\n"
        "- PDF导出接口"
    )
    st.markdown(
        "**本地AI排查命令：**\n"
        "```powershell\n"
        "ollama pull qwen2.5:3b\n"
        "ollama serve\n"
        "ollama run qwen2.5:3b\n"
        "```"
    )

ollama_ok, ollama_msg = check_ollama_status()
if mode == "规则+AI增强（推荐）":
    if ollama_ok:
        st.success(ollama_msg)
    else:
        st.warning(f"{ollama_msg}。当前仍可切换到“仅规则兜底”生成可编辑JSON和Word。")

uploaded_protocol = st.file_uploader("上传方案（PDF / DOCX / TXT）", type=["pdf", "docx", "txt"])
uploaded_template = st.file_uploader("上传Word模板（可选）", type=["docx"])

template_path = DEFAULT_TEMPLATE
if uploaded_template:
    template_path = OUTPUT_DIR / uploaded_template.name
    template_path.write_bytes(uploaded_template.getvalue())

st.info(f"当前模板：{template_path.name}")

if uploaded_protocol:
    raw_text = extract_text_from_file(uploaded_protocol)
    st.subheader("协议文本预览")
    st.text_area("提取结果", raw_text[:20000], height=260)

    if raw_text.startswith("文件解析失败") or raw_text.startswith("PDF解析失败"):
        st.error(raw_text)
    else:
        if st.button("AI生成结构化JSON"):
            use_ai = mode == "规则+AI增强（推荐）" and ollama_ok
            if mode == "规则+AI增强（推荐）" and not ollama_ok:
                st.warning("本地AI不可用，系统已自动切换为规则兜底模式生成。")
            with st.spinner("系统正在解析协议并生成结构化JSON，请稍候..."):
                data = build_v62_plan_json(raw_text, use_ai=use_ai)
            if isinstance(data, dict) and data.get("_error"):
                st.error(data["_error"])
                if data.get("_detail"):
                    st.code(str(data["_detail"]))
                st.info("建议切换到“仅规则兜底”，或确认 Ollama/qwen2.5:3b 是否可用。")
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
        audit_type = st.text_input("稽查类型", value=data.get("project", {}).get("audit_type", "中心常规稽查"))
    with c3:
        author = st.text_input("撰写人", value=data.get("AUTHOR", "苗田"))
        approver = st.text_input("审批人", value=data.get("APPROVER", "张艳"))
        audit_company = st.text_input("稽查公司", value=data.get("AUDIT_COMPANY", "北京万宁睿和医药科技有限公司"))

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

    edited_text = st.text_area(
        "JSON预览与编辑",
        value=json.dumps(data, ensure_ascii=False, indent=2),
        height=560,
    )

    tabs = st.tabs(["项目概览", "入排标准", "流程要求", "AI稽查重点", "RBQM", "访谈问题", "发现/CAPA草案", "完整JSON"])
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
            docx_path = OUTPUT_DIR / "PlanGen_v63_Output.docx"
            generate_docx_from_template(template_path=template_path, data=edited_data, output_path=docx_path)
            st.success("DOCX 生成成功")
            st.download_button("下载 DOCX", data=docx_path.read_bytes(), file_name=docx_path.name)
            pdf_path = convert_docx_to_pdf(docx_path, output_dir=OUTPUT_DIR)
            if pdf_path and Path(pdf_path).exists():
                st.download_button("下载 PDF", data=Path(pdf_path).read_bytes(), file_name=Path(pdf_path).name)
            else:
                st.info("未生成 PDF。若需要 PDF，请先在本机安装 LibreOffice。")
        except Exception as e:
            st.error(f"正式生成文档失败：{e}")
