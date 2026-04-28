from __future__ import annotations

import re
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, List

import docx
try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

from .ai_engine import call_ai_json, call_ai_text

DIRTY_FIELD_WORDS = ["版本号", "版本日期", "保密", "SUSAR", "SAE", "统计方法", "页", "数据库锁定", "严重不良事件", "安全性"]
SPONSOR_SUFFIXES = ["有限公司", "股份有限公司", "有限责任公司", "科技有限公司", "药业有限公司", "医药科技有限公司"]

def _save_uploaded_temp_file(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix.lower()
    temp_name = f"{uuid.uuid4().hex}{suffix}"
    tmp = Path(tempfile.gettempdir()) / temp_name
    tmp.write_bytes(uploaded_file.getvalue())
    return tmp

def extract_text_from_file(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix.lower()
    tmp = _save_uploaded_temp_file(uploaded_file)
    try:
        if suffix == ".txt":
            return tmp.read_text(encoding="utf-8", errors="ignore")
        if suffix == ".docx":
            doc = docx.Document(str(tmp))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if suffix == ".pdf":
            if PdfReader is None:
                return "当前环境未安装 PyPDF2，无法解析 PDF。"
            reader = PdfReader(str(tmp))
            texts = []
            for page in reader.pages:
                try:
                    texts.append(page.extract_text() or "")
                except Exception:
                    pass
            return "\n".join(texts)
        return ""
    except Exception as e:
        return f"文件解析失败：{e}"

def normalize_text(text: str) -> str:
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text

def _clean_line(s: str) -> str:
    s = normalize_text(s).strip("•·- \t")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def _sanitize_single_line_field(value: str, max_len: int = 120) -> str:
    if not value:
        return ""
    value = value.replace("\n", " ").replace("\r", " ")
    value = _clean_line(value)
    return value[:max_len]

def _looks_dirty(value: str) -> bool:
    if not value:
        return True
    if len(value) > 160:
        return True
    hit = sum(1 for w in DIRTY_FIELD_WORDS if w in value)
    return hit >= 2

def _extract_project_name_rule(text: str) -> str:
    for p in [r"项目名称[:：]\s*(.+)", r"研究名称[:：]\s*(.+)", r"题目[:：]\s*(.+)"]:
        m = re.search(p, text, re.I)
        if m:
            v = _sanitize_single_line_field(m.group(1), 160)
            if not _looks_dirty(v):
                return v
    for line in text.splitlines():
        line = _clean_line(line)
        if "临床研究方案" in line and len(line) < 160:
            return line
    return "待补充项目名称"

def _extract_sponsor_rule(text: str) -> str:
    m = re.search(r"申办方[:：]\s*([^\n\r]{2,100})", text)
    if m:
        candidate = _sanitize_single_line_field(m.group(1), 100)
        if not _looks_dirty(candidate):
            return candidate
    header_text = "\n".join(text.splitlines()[:100])
    cands = []
    for suffix in SPONSOR_SUFFIXES:
        pattern = rf"([^\n\r，,；;（）(){{}}]{{2,50}}?{re.escape(suffix)})"
        for mm in re.finditer(pattern, header_text):
            c = _sanitize_single_line_field(mm.group(1), 100)
            if c and not _looks_dirty(c):
                cands.append(c)
    if cands:
        cands = sorted(set(cands), key=len)
        return cands[0]
    return "待补充申办方"

def _extract_protocol_code_rule(text: str) -> str:
    m = re.search(r"\b([A-Z]{2,}[A-Z0-9-]{3,})\b", text)
    return _clean_line(m.group(1)) if m else ""

def _extract_version_date_rule(text: str):
    version = ""
    date_ = ""
    m = re.search(r"版本号[:：]?\s*(V?\d+(?:\.\d+)*)", text, re.I)
    if m:
        version = _clean_line(m.group(1))
    m = re.search(r"版本日期[:：]?\s*([0-9]{4}[年/-][0-9]{1,2}[月/-][0-9]{1,2}日?)", text, re.I)
    if m:
        date_ = _clean_line(m.group(1))
    return version or "V1.0", date_

def _extract_indication_rule(text: str) -> str:
    m = re.search(r"适应症[:：]\s*(.+)", text)
    if m:
        v = _sanitize_single_line_field(m.group(1), 100)
        if not _looks_dirty(v):
            return v
    title = _extract_project_name_rule(text)
    m = re.search(r"治疗(.+?)的有效性和安全性", title)
    if m:
        return _sanitize_single_line_field(m.group(1), 100)
    return "待补充适应症"

def _infer_phase_rule(text: str) -> str:
    m = re.search(r"(I期|II期|III期|IV期|Ⅰ期|Ⅱ期|Ⅲ期|Ⅳ期|\bPhase\s*[1-4IVX]+\b)", text, re.I)
    return m.group(1) if m else "待补充研究阶段"

def _extract_between(text: str, start_keywords: List[str], end_keywords: List[str]) -> str:
    lines = text.splitlines()
    start_idx = None
    for i, ln in enumerate(lines):
        if any(k in ln for k in start_keywords):
            start_idx = i + 1
            break
    if start_idx is None:
        return ""
    end_idx = len(lines)
    for j in range(start_idx, len(lines)):
        if any(k in lines[j] for k in end_keywords):
            end_idx = j
            break
    return "\n".join(lines[start_idx:end_idx]).strip()

def _extract_study_design_rule(text: str) -> str:
    design = _extract_between(text, ["研究设计", "试验设计", "Study Design"], ["入选标准", "纳入标准", "排除标准", "主要终点", "研究目的"])
    if design:
        return _sanitize_single_line_field(design, 400)
    m = re.search(r"(一项.*?研究.*?)(?:\n|$)", text, re.S)
    if m:
        return _sanitize_single_line_field(m.group(1), 400)
    return "待补充研究设计"

def _extract_primary_endpoint_rule(text: str) -> str:
    block = _extract_between(text, ["主要终点", "Primary Endpoint", "研究目的和终点"], ["次要终点", "安全性", "不良事件", "试验用药品", "生物样本"])
    lines = [_clean_line(x) for x in block.splitlines() if _clean_line(x)]
    lines = [x for x in lines if "主要终点" not in x and "Primary Endpoint" not in x]
    val = "；".join(lines[:3]) if lines else "待补充主要终点"
    return _sanitize_single_line_field(val, 400)

def _split_text_for_ai(text: str, max_chars: int = 6000) -> str:
    return normalize_text(text)[:max_chars]

def _ai_extract_basic(text: str) -> dict:
    prompt = f"""
你是临床试验方案结构化抽取助手。请从以下方案文本中提取基础字段，只返回JSON，不要解释。

字段：
- project_title
- sponsor_name
- protocol_code
- version_no
- version_date
- indication
- phase
- study_design
- primary_endpoint

要求：
1. sponsor_name 必须是公司全称
2. 不要输出页码、保密、版本历史等污染内容
3. study_design 和 primary_endpoint 尽量简洁
4. 不确定则返回空字符串

文本：
{_split_text_for_ai(text, 6000)}
"""
    return call_ai_json(prompt, timeout=45)

def _ai_extract_items_with_focus(text: str) -> dict:
    prompt = f"""
你是临床试验方案结构化抽取助手。请从以下方案文本中提取条目，并同时给出“重点关注”。只返回JSON。

提取这些部分：
- inclusion_items（入选标准）
- exclusion_items（排除标准）
- process_requirements（流程要求，type 只能是：randomization / rescreen / stop_treatment / withdrawal / prohibited_med / dose_adjustment）

要求：
1. 每条拆成独立条目
2. 尽量保留原编号，如 2a, 2b(i), 13a
3. focus 必须具体，指出要查什么记录/文件/系统，并关注阈值、时间窗、一致性、研究者判断
4. 不要混入页眉页脚、版本历史、缩略语
5. 若条目很多，优先返回最关键的前 12 条入选、前 18 条排除、前 6 条流程要求

格式：
{{
  "inclusion_items":[{{"id":"","text":"","focus":""}}],
  "exclusion_items":[{{"id":"","text":"","focus":""}}],
  "process_requirements":[{{"type":"","text":"","focus":""}}]
}}

文本：
{_split_text_for_ai(text, 8000)}
"""
    return call_ai_json(prompt, timeout=60)

def _generate_rbqm_ai(context: str) -> str:
    prompt = f"""
你是临床试验 RBQM 专家。请根据以下协议内容，生成 5 条 RBQM 风险导向质控策略。
要求：
1. 聚焦高/中/低风险中心和受试者抽查
2. 聚焦主要终点、关键实验室、数据溯源、试验药物链条
3. 只输出纯文本，每行1条，不要JSON

协议摘要：
{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)

def _generate_interview_questions_ai(context: str) -> str:
    prompt = f"""
你是临床试验稽查专家。请基于以下方案内容，为申办方/研究者生成 8 条访谈问题。
要求：
1. 问题具体、可现场使用
2. 覆盖入排、随机、终止/退出、用药、安全性、数据溯源
3. 只输出纯文本，每行1条

协议摘要：
{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)

def _generate_finding_capa_ai(context: str) -> str:
    prompt = f"""
你是临床试验质控专家。请基于以下方案内容，生成：
1. 常见发现方向
2. 对应 CAPA 草案建议

要求：
- 输出 6 条
- 每条包含“发现方向：...；CAPA建议：...”
- 只输出纯文本，每行1条

协议摘要：
{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)

def _fallback_rbqm(criteria_rows, exclusion_rows):
    return "\n".join([
        "1. 高风险中心：100%审核受试者文件夹，重点核查 SAE、主要终点、关键实验室和 IMP 链条",
        "2. 中风险中心：50%抽查，优先核查首例/末例、异常值、方案偏离和数据修改频繁病例",
        "3. 低风险中心：30%抽查，重点核查知情同意、关键访视时间窗和源数据可溯源性",
        "4. 高风险受试者优先：SAE病例、失访病例、关键实验室异常、主要终点相关病例",
        f"5. 本项目入排标准复杂度参考：入选{min(len(criteria_rows),12)}项，排除{min(len(exclusion_rows),18)}项",
    ])

def _fallback_interview_questions():
    return "\n".join([
        "1. 本研究受试者入组时，如何确认关键入排标准及时间窗满足方案要求？",
        "2. 随机化是否在规定时点完成，如何核对随机号与药物编号绑定关系？",
        "3. 如发生复筛，研究团队如何判断哪些检查可以豁免？",
        "4. 当受试者出现终止治疗或退出研究情况时，后续安全性随访如何执行？",
        "5. 禁止合并用药和限制用药由谁审核，如何留痕？",
        "6. 关键实验室和主要终点数据如何保证源数据与 EDC 一致？",
        "7. AE/SAE 的识别、分级、因果性判断和上报链条如何执行？",
        "8. IMP 接收、储存、分发、使用、回收如何实现账物一致？",
    ])

def _fallback_finding_capa():
    return "\n".join([
        "1. 发现方向：入排标准核查证据不足；CAPA建议：补充原始记录核查要点，完善筛选留痕与研究者确认。",
        "2. 发现方向：随机化时点或分层因素记录不完整；CAPA建议：建立 IWRS 核对清单并纳入稽查前复核。",
        "3. 发现方向：关键检查未严格满足时间窗；CAPA建议：建立访视时间窗预警与偏差升级机制。",
        "4. 发现方向：禁用药或合并治疗审核不充分；CAPA建议：强化药物审核节点并留存医嘱/药房双重核对记录。",
        "5. 发现方向：AE/SAE 识别或随访记录不足；CAPA建议：补强安全性判定培训并完善随访闭环。",
        "6. 发现方向：源数据与 EDC 一致性不足；CAPA建议：对关键数据点执行重点 SDV 与系统溯源复核。",
    ])

def build_v62_plan_json(text: str, use_ai: bool = True) -> Dict[str, Any]:
    text = normalize_text(text)

    project_name = _extract_project_name_rule(text)
    sponsor = _extract_sponsor_rule(text)
    protocol_code = _extract_protocol_code_rule(text)
    version_no, version_date = _extract_version_date_rule(text)
    indication = _extract_indication_rule(text)
    phase = _infer_phase_rule(text)
    study_design = _extract_study_design_rule(text)
    primary_endpoint = _extract_primary_endpoint_rule(text)

    criteria_rows = [{"criterion":"待补充", "ai_focus":"待补充"}]
    exclusion_rows = [{"criterion":"待补充", "ai_focus":"待补充"}]
    process_rows = [{"requirement":"待补充", "focus":"待补充"}]

    if use_ai:
        basic_ai = _ai_extract_basic(text)
        if isinstance(basic_ai, dict) and basic_ai.get("_error"):
            return basic_ai

        p = _sanitize_single_line_field(basic_ai.get("project_title",""), 160)
        s = _sanitize_single_line_field(basic_ai.get("sponsor_name",""), 100)
        c = _sanitize_single_line_field(basic_ai.get("protocol_code",""), 50)
        v = _sanitize_single_line_field(basic_ai.get("version_no",""), 30)
        d = _sanitize_single_line_field(basic_ai.get("version_date",""), 40)
        ind = _sanitize_single_line_field(basic_ai.get("indication",""), 100)
        ph = _sanitize_single_line_field(basic_ai.get("phase",""), 30)
        sd = _sanitize_single_line_field(basic_ai.get("study_design",""), 400)
        pe = _sanitize_single_line_field(basic_ai.get("primary_endpoint",""), 400)

        if p and not _looks_dirty(p): project_name = p
        if s and not _looks_dirty(s): sponsor = s
        if c and not _looks_dirty(c): protocol_code = c
        if v: version_no = v
        if d: version_date = d
        if ind and not _looks_dirty(ind): indication = ind
        if ph: phase = ph
        if sd and not _looks_dirty(sd): study_design = sd
        if pe and not _looks_dirty(pe): primary_endpoint = pe

        items_ai = _ai_extract_items_with_focus(text)
        if isinstance(items_ai, dict) and items_ai.get("_error"):
            return items_ai

        inclusion_items = items_ai.get("inclusion_items", []) or []
        exclusion_items = items_ai.get("exclusion_items", []) or []
        process_items = items_ai.get("process_requirements", []) or []

        if inclusion_items:
            criteria_rows = []
            for item in inclusion_items[:12]:
                item_id = _sanitize_single_line_field(str(item.get("id","")), 20)
                item_text = _sanitize_single_line_field(str(item.get("text","")), 400)
                focus = _sanitize_single_line_field(str(item.get("focus","")), 400)
                if not item_text:
                    continue
                criterion = f"入选标准{item_id}：{item_text}" if item_id else f"入选标准：{item_text}"
                criteria_rows.append({"criterion": criterion, "ai_focus": focus or "待补充"})
            if not criteria_rows:
                criteria_rows = [{"criterion":"待补充", "ai_focus":"待补充"}]

        if exclusion_items:
            exclusion_rows = []
            for item in exclusion_items[:18]:
                item_id = _sanitize_single_line_field(str(item.get("id","")), 20)
                item_text = _sanitize_single_line_field(str(item.get("text","")), 400)
                focus = _sanitize_single_line_field(str(item.get("focus","")), 400)
                if not item_text:
                    continue
                criterion = f"排除标准{item_id}：{item_text}" if item_id else f"排除标准：{item_text}"
                exclusion_rows.append({"criterion": criterion, "ai_focus": focus or "待补充"})
            if not exclusion_rows:
                exclusion_rows = [{"criterion":"待补充", "ai_focus":"待补充"}]

        if process_items:
            label_map = {
                "randomization": "随机化要求",
                "rescreen": "复筛要求",
                "stop_treatment": "终止治疗要求",
                "withdrawal": "退出研究要求",
                "prohibited_med": "禁止合并治疗要求",
                "dose_adjustment": "药物剂量调整规则",
            }
            process_rows = []
            for item in process_items[:6]:
                typ = str(item.get("type","")).strip()
                txt = _sanitize_single_line_field(str(item.get("text","")), 450)
                focus = _sanitize_single_line_field(str(item.get("focus","")), 400)
                if not txt:
                    continue
                label = label_map.get(typ, "流程要求")
                process_rows.append({"requirement": f"{label}：{txt}", "focus": focus or "待补充"})
            if not process_rows:
                process_rows = [{"requirement":"待补充", "focus":"待补充"}]

    ai_points = []
    for x in criteria_rows[:4]:
        ai_points.append(f"入选标准核查：{x['criterion']} → {x['ai_focus']}")
    for x in exclusion_rows[:4]:
        ai_points.append(f"排除标准核查：{x['criterion']} → {x['ai_focus']}")
    if primary_endpoint and primary_endpoint != "待补充主要终点":
        ai_points.append(f"主要终点核查：{primary_endpoint} → 重点关注关键时间窗、源数据溯源和评估一致性")
    for x in process_rows[:4]:
        ai_points.append(f"流程要求核查：{x['requirement']} → {x['focus']}")
    ai_points += [
        "知情同意核查：重点关注版本、签署时间、执行人员授权及重签情况",
        "安全性核查：重点关注 AE/SAE 识别、分级、因果性判断、上报时效与随访",
        "IMP/试验药物核查：重点关注接收、储存、分发、使用、回收及账物一致性",
        "数据溯源核查：重点关注 HIS/LIS/PACS/EDC 与原始记录一致性",
    ]

    if use_ai:
        rbqm_txt = _generate_rbqm_ai(text)
        interview_txt = _generate_interview_questions_ai(text)
        finding_capa_txt = _generate_finding_capa_ai(text)
        if rbqm_txt.startswith("AI调用失败"): rbqm_txt = _fallback_rbqm(criteria_rows, exclusion_rows)
        if interview_txt.startswith("AI调用失败"): interview_txt = _fallback_interview_questions()
        if finding_capa_txt.startswith("AI调用失败"): finding_capa_txt = _fallback_finding_capa()
    else:
        rbqm_txt = _fallback_rbqm(criteria_rows, exclusion_rows)
        interview_txt = _fallback_interview_questions()
        finding_capa_txt = _fallback_finding_capa()

    return {
        "AUTHOR": "苗田",
        "APPROVER": "张艳",
        "VERSION_NO": version_no,
        "VERSION_DATE": version_date or "",
        "AUDIT_TYPE": "中心常规稽查",
        "AUDIT_COMPANY": "北京万宁睿和医药科技有限公司",
        "PROJECT_TITLE": project_name,
        "PROJECT_CODE": protocol_code,
        "SPONSOR_NAME": sponsor,
        "project": {
            "name": project_name,
            "sponsor": sponsor,
            "indication": indication,
            "phase": phase,
            "version_no": version_no,
            "version_date": version_date or "",
            "audit_type": "中心常规稽查",
            "protocol_code": protocol_code,
        },
        "protocol_analysis": {
            "study_design": study_design,
            "primary_endpoint": primary_endpoint,
            "key_criteria": "；".join([x["criterion"] for x in criteria_rows[:5]]) if criteria_rows else "待补充关键纳排标准",
        },
        "ai_audit_key_points": "\n".join(f"{i+1}. {x}" for i, x in enumerate(ai_points)),
        "rbqm_strategy": rbqm_txt,
        "interview_questions": interview_txt,
        "finding_capa_draft": finding_capa_txt,
        "sampling_strategy": "中心抽查比例建议：高风险100%，中风险50%，低风险30%；首例/末例、SAE病例、异常数据病例优先抽查。",
        "capa_classification": "\n".join([
            "1. Critical：严重影响受试者权益/安全或数据真实性、完整性的问题",
            "2. Major：重要不依从，可能影响关键数据质量或研究执行质量的问题",
            "3. Minor：一般性偏差，对整体影响较小但需整改的问题",
        ]),
        "audit_plan_summary": "基于 Protocol 自动生成的稽查计划草稿，建议人工复核研究设计、终点、关键时间窗、IMP管理及安全性要求。",
        "criteria_ai_rows": criteria_rows,
        "exclusion_ai_rows": exclusion_rows,
        "risk_analysis_rows": [
            {"risk_factor": "知情同意", "detail": "重点核查版本、签署时序、授权人员、重签和病历记录。"},
            {"risk_factor": "安全性", "detail": "重点核查 AE/SAE 识别、分级、因果性判断、上报时效与随访。"},
            {"risk_factor": "数据溯源", "detail": "重点核查 HIS/LIS/PACS/EDC 与原始记录一致性。"},
            {"risk_factor": "IMP管理", "detail": "重点核查接收、储存、分发、使用、回收和账物一致性。"},
        ],
        "subject_rows": [{"subject_id": "待填写", "protocol_version": "待填写", "summary": "建议优先录入首例/末例/SAE病例/偏差病例。"}],
        "assignment_rows": [
            {"seq": "1", "process": "研究者文件夹/受试者文件夹核查", "assignee": "待填写", "plan_time": "待填写"},
            {"seq": "2", "process": "安全性信息核查", "assignee": "待填写", "plan_time": "待填写"},
            {"seq": "3", "process": "IMP/样本/系统溯源核查", "assignee": "待填写", "plan_time": "待填写"},
        ],
        "process_requirement_rows": process_rows,
        "secondary_endpoint_rows": [{"objective": "安全性评估", "endpoint": "AE/SAE、实验室、生命体征、ECG/影像等。"}],
        "safety_focus_rows": [
            {"category": "生命体征/护理记录", "focus": "核对流程、异常值判断与原始记录完整性。"},
            {"category": "不良事件/严重不良事件", "focus": "核对识别、分级、因果性、上报链条与随访。"},
        ],
        "defect_rows": [
            {"category": "知情同意书（ICF）的签署和记录", "yes": "", "no": "√", "minor": "", "major": ""},
            {"category": "安全性信息评估、记录与报告", "yes": "", "no": "√", "minor": "", "major": ""},
            {"category": "试验用药品管理", "yes": "", "no": "√", "minor": "", "major": ""},
        ],
        "report_send_rows": [{"name": "待填写", "email": "待填写", "title_company": "待填写"}],
    }
