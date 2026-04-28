from __future__ import annotations

import re
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, List, Tuple

import docx
try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

from .ai_engine import call_ai_json, call_ai_text

DIRTY_FIELD_WORDS = ["版本号", "版本日期", "保密", "SUSAR", "SAE", "统计方法", "页", "数据库锁定", "严重不良事件", "安全性"]
SPONSOR_SUFFIXES = ["有限公司", "股份有限公司", "有限责任公司", "科技有限公司", "药业有限公司", "医药科技有限公司"]

# 行业级重点关注规则：当本地AI不稳定或输出太泛时，用规则生成专业核查动作。
FOCUS_RULES = [
    (["年龄", "岁"], "查阅身份证/病历首页/知情同意书签署日期，确认签署时年龄符合方案范围；核对筛选表、eCRF人口学信息与源文件一致。"),
    (["知情同意", "自愿"], "查阅ICF原件、伦理批准版本和签署记录，确认签署日期早于任何筛选程序；核对签署人、执行知情同意人员授权、版本号及受试者留存副本。"),
    (["妊娠", "哺乳", "生育", "避孕"], "查阅妊娠检测报告、病历及访视记录，确认检测时间窗符合方案；核对受试者避孕承诺/生育意向声明及随访期间避孕措施执行记录。"),
    (["诊断", "疾病", "适应症", "溃疡性结肠炎", "克罗恩", "肝病", "肿瘤"], "查阅病历、诊断证明、内镜/影像/病理报告等源文件，确认诊断依据充分；核对诊断时间、疾病类型、疾病活动度与筛选记录/eCRF一致。"),
    (["纤维蛋白原", "血小板", "血红蛋白", "肌酐", "HIV", "梅毒", "实验室", "检测", "阴性", "阳性", "CRP", "PCT", "D-二聚体", "FDP", "APTT", "PT", "TT", "INR"], "查阅LIS/中心实验室原始报告，核对采样时间、检测结果、单位、阈值和方案时间窗；确认异常值研究者已评估，并比对eCRF录入与源数据一致。"),
    (["Child-Pugh", "ECOG", "NYHA", "West-Haven", "评分", "分级", "Mayo", "CDAI"], "查阅评分量表、组成项原始记录及研究者判断依据，复核评分计算过程和分级结果；确认评分时间在筛选期/规定访视窗内，并与eCRF记录一致。"),
    (["过敏", "辅料", "不能耐受", "超敏"], "查阅既往史、过敏史、用药史及筛选问诊记录，确认无相关过敏或不能耐受；核对受试者自述、病历记录和研究者判断一致。"),
    (["感染", "DIC", "纤溶", "出血", "血栓", "栓塞", "脑卒中", "心血管", "心肌梗死", "心律失常"], "查阅病历、医嘱、实验室检查、影像/PACS及专科评估记录，确认未触发排除条件；核对事件发生时间、严重程度、研究者判断和筛选时间窗。"),
    (["恶性肿瘤", "精神", "癫痫", "认知", "吸毒", "酗酒"], "查阅既往史、专科评估、用药记录和研究者筛选判断，确认无影响依从性/安全性或符合排除标准的情况；核对评估结论有研究者签名/日期。"),
    (["华法林", "阿司匹林", "肝素", "凝血因子", "氨甲环酸", "维生素K", "血凝酶", "合并治疗", "禁用药", "限制用药"], "查阅合并用药、医嘱、药房记录及受试者用药问诊，核对禁用/限制药物洗脱期、使用时间与随机/给药时间关系；如违规，确认记录为方案偏离并完成医学评估。"),
    (["随机", "IWRS", "分层", "1:1", "随机化"], "查阅IWRS/IRT随机记录、筛选合格确认、随机号与药物编号绑定记录，确认随机在方案规定时点完成；核对分层因素选择正确并与源数据/eCRF一致。"),
    (["复筛", "重新筛选", "28天", "豁免"], "查阅ICF签署日期、筛选检查日期和首次给药/随机日期，确认是否超出复筛触发条件；对豁免项目核对近期合格结果、时间窗和研究者书面确认。"),
    (["终止治疗", "退出研究", "失访", "撤回"], "查阅终止/退出记录、AE/SAE、疗效不佳或受试者撤回原因说明，确认研究者判断依据充分；核对退出前后安全性检查、随访安排和eCRF状态。"),
    (["剂量", "给药", "输注", "体重", "目标值", "用药前"], "查阅给药医嘱、体重记录、剂量计算表、给药/输注记录和给药前关键检测结果，复核公式、目标值、给药间隔和实际剂量；核对药物发放使用记录与eCRF一致。"),
]


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
    s = normalize_text(str(s)).strip("•·- \t")
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
    if len(value) > 180:
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
        return sorted(set(cands), key=len)[0]
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


def _section(text: str, starts: List[str], ends: List[str], max_chars: int = 6000) -> str:
    block = _extract_between(text, starts, ends)
    return normalize_text(block)[:max_chars]


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


def _split_items_rule(block: str, prefix: str, limit: int) -> List[Dict[str, str]]:
    block = normalize_text(block)
    raw = [_clean_line(x) for x in block.splitlines() if _clean_line(x)]
    items, buf = [], ""

    def flush():
        nonlocal buf
        if buf and len(buf) >= 5:
            items.append(buf)
        buf = ""

    for ln in raw:
        if re.match(r"^(\d+[a-zA-Z]?(?:\([ivx]+\))?[\.、)）:：]|[a-zA-Z][\.、)）:：])", ln, re.I):
            flush()
            buf = ln
        elif buf:
            buf += " " + ln
        else:
            buf = ln
    flush()

    rows = []
    seen = set()
    for it in items[:limit]:
        it = _clean_line(it)
        if it in seen:
            continue
        seen.add(it)
        rows.append({"criterion": f"{prefix}{it}", "ai_focus": _build_professional_focus(it)})
    return rows


def _build_professional_focus(item_text: str) -> str:
    text = _clean_line(item_text)
    hits = []
    for keywords, focus in FOCUS_RULES:
        if any(k.lower() in text.lower() for k in keywords):
            hits.append(focus)
    if not hits:
        hits.append("查阅病历、原始记录、实验室/影像检查、合并用药、研究者判断及EDC录入，确认方案要求、时间窗、阈值和源数据一致性均满足要求。")
    # 合并最多两条，避免过长，但保证专业性
    return "；".join(hits[:2])


def _focus_is_weak(focus: str) -> bool:
    if not focus or focus == "待补充":
        return True
    weak_words = ["核查一致性", "确认符合", "检查相关记录", "核查资料", "待补充"]
    return len(focus) < 35 or any(w in focus for w in weak_words)


def _ai_extract_basic(text: str) -> dict:
    prompt = f"""
你是临床试验方案结构化抽取助手。请从以下方案文本中提取基础字段，只返回JSON，不要解释。
字段：project_title, sponsor_name, protocol_code, version_no, version_date, indication, phase, study_design, primary_endpoint。
要求：sponsor_name必须是公司全称；不要输出页码、保密、版本历史等污染内容；不确定则返回空字符串。
文本：
{normalize_text(text)[:6000]}
"""
    return call_ai_json(prompt, timeout=45)


def _ai_extract_items_with_focus(text: str) -> dict:
    inclusion = _section(text, ["入选标准", "纳入标准", "Inclusion Criteria"], ["排除标准", "Exclusion Criteria", "随机", "研究目的", "主要终点"], 5000)
    exclusion = _section(text, ["排除标准", "Exclusion Criteria"], ["随机", "随机化", "研究目的", "主要终点", "安全性", "终止治疗"], 6000)
    process = _section(text, ["随机", "随机化", "复筛", "终止治疗", "退出研究", "禁止合并治疗", "剂量"], ["安全性", "研究目的", "主要终点", "统计"], 4000)

    prompt = f"""
你是资深GCP稽查专家。请基于下面的协议片段抽取条目，并生成“行业级重点关注”。只返回JSON。

必须输出：
{{
  "inclusion_items":[{{"id":"","text":"","focus":""}}],
  "exclusion_items":[{{"id":"","text":"","focus":""}}],
  "process_requirements":[{{"type":"randomization/rescreen/stop_treatment/withdrawal/prohibited_med/dose_adjustment","text":"","focus":""}}]
}}

重点关注focus必须是“核查动作”，不得泛泛而谈，至少覆盖以下要素中的3类：
1. 需要查阅的源文件/系统：病历、医嘱、HIS、LIS、PACS、EDC/eCRF、IWRS/IRT、药房/IMP记录。
2. 时间窗：筛选期、首次用药前、随机前、给药前、访视窗等。
3. 阈值/分级/评分：实验室阈值、评分计算、分级标准、诊断依据。
4. 一致性：源数据与eCRF/EDC一致、系统间一致、研究者判断与记录一致。
5. 研究者责任：研究者判断、签名日期、医学评估、方案偏离记录。

入选标准片段：
{inclusion}

排除标准片段：
{exclusion}

流程要求片段：
{process}
"""
    return call_ai_json(prompt, timeout=60)


def _generate_rbqm_ai(context: str) -> str:
    prompt = f"""
你是临床试验 RBQM 专家。请生成5条风险导向稽查策略，每条必须包含风险对象、核查动作、优先级逻辑。只输出纯文本，每行1条。
协议摘要：{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)


def _generate_interview_questions_ai(context: str) -> str:
    prompt = f"""
你是临床试验稽查专家。请生成8条现场访谈问题，覆盖PI/研究医生/CRC/药品管理员。问题要能直接现场提问。只输出纯文本，每行1条。
协议摘要：{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)


def _generate_finding_capa_ai(context: str) -> str:
    prompt = f"""
你是临床试验质控专家。请生成6条常见发现方向和CAPA草案。每条格式：发现方向：...；CAPA建议：...。只输出纯文本。
协议摘要：{context[:4000]}
"""
    return call_ai_text(prompt, timeout=35)


def _fallback_rbqm(criteria_rows, exclusion_rows):
    return "\n".join([
        "1. 高风险中心：优先选择入组速度快、SAE/方案偏离较多、关键数据质疑率高或新启动中心，执行100%受试者文件夹核查。",
        "2. 关键数据核查：主要终点、关键实验室、评分量表、随机化和给药记录执行重点SDV，并核对HIS/LIS/PACS/EDC一致性。",
        "3. 高风险受试者优先：首例/末例、SAE病例、退出/失访病例、关键实验室异常、方案偏离多发病例优先抽查。",
        "4. IMP链条核查：重点核查接收、储存、温控、分发、使用、回收、销毁和账物一致，发现异常需追踪偏差和CAPA。",
        f"5. 入排复杂度提示：当前入选{min(len(criteria_rows),12)}项、排除{min(len(exclusion_rows),18)}项，应将时间窗、阈值、评分和研究者判断作为稽查重点。",
    ])


def _fallback_interview_questions():
    return "\n".join([
        "1. PI如何确认受试者在签署ICF前未执行任何筛选程序？",
        "2. 研究医生如何核对关键入排标准的诊断依据、实验室阈值和时间窗？",
        "3. 随机化前由谁确认受试者筛选合格，IWRS/IRT记录如何留存？",
        "4. 如发生复筛或检查豁免，研究者如何书面判断并留痕？",
        "5. 禁止/限制合并用药由谁审核，如何与医嘱、药房记录和eCRF核对？",
        "6. AE/SAE如何识别、分级、判断因果性并完成上报和随访？",
        "7. 关键实验室、评分量表和主要终点数据如何保证源数据与EDC一致？",
        "8. IMP接收、储存、分发、使用、回收全过程如何确保账物一致和温控合规？",
    ])


def _fallback_finding_capa():
    return "\n".join([
        "1. 发现方向：入排标准核查证据不足；CAPA建议：建立筛选核查清单，补充病历、实验室、影像和研究者判断留痕。",
        "2. 发现方向：关键检查时间窗不符合方案；CAPA建议：设置访视/采样时间窗预警，超窗时记录方案偏离并完成医学评估。",
        "3. 发现方向：EDC与源数据不一致；CAPA建议：对关键数据点执行二次SDV，明确更正理由、签名和日期。",
        "4. 发现方向：随机化或分层因素记录不完整；CAPA建议：补充IWRS/IRT截图和筛选合格确认记录，培训随机化流程。",
        "5. 发现方向：禁用药审核不足；CAPA建议：建立合并用药审核机制，核对医嘱、药房、受试者问诊和eCRF记录。",
        "6. 发现方向：AE/SAE随访闭环不足；CAPA建议：完善安全性事件追踪表，明确责任人和上报/随访时限。",
    ])


def _rows_from_ai_or_rule(text: str, items_ai: Dict[str, Any]) -> Tuple[List[Dict[str, str]], List[Dict[str, str]], List[Dict[str, str]]]:
    inclusion_items = items_ai.get("inclusion_items", []) if isinstance(items_ai, dict) else []
    exclusion_items = items_ai.get("exclusion_items", []) if isinstance(items_ai, dict) else []
    process_items = items_ai.get("process_requirements", []) if isinstance(items_ai, dict) else []

    criteria_rows = []
    for item in inclusion_items[:12]:
        item_id = _sanitize_single_line_field(str(item.get("id", "")), 20)
        item_text = _sanitize_single_line_field(str(item.get("text", "")), 420)
        focus = _sanitize_single_line_field(str(item.get("focus", "")), 520)
        if not item_text:
            continue
        if _focus_is_weak(focus):
            focus = _build_professional_focus(item_text)
        criterion = f"入选标准{item_id}：{item_text}" if item_id else f"入选标准：{item_text}"
        criteria_rows.append({"criterion": criterion, "ai_focus": focus})

    exclusion_rows = []
    for item in exclusion_items[:18]:
        item_id = _sanitize_single_line_field(str(item.get("id", "")), 20)
        item_text = _sanitize_single_line_field(str(item.get("text", "")), 420)
        focus = _sanitize_single_line_field(str(item.get("focus", "")), 520)
        if not item_text:
            continue
        if _focus_is_weak(focus):
            focus = _build_professional_focus(item_text)
        criterion = f"排除标准{item_id}：{item_text}" if item_id else f"排除标准：{item_text}"
        exclusion_rows.append({"criterion": criterion, "ai_focus": focus})

    if not criteria_rows:
        inc_block = _section(text, ["入选标准", "纳入标准"], ["排除标准", "随机", "研究目的"], 5000)
        criteria_rows = _split_items_rule(inc_block, "入选标准", 12) or [{"criterion": "待补充", "ai_focus": "待补充"}]
    if not exclusion_rows:
        exc_block = _section(text, ["排除标准"], ["随机", "随机化", "研究目的", "主要终点"], 6000)
        exclusion_rows = _split_items_rule(exc_block, "排除标准", 18) or [{"criterion": "待补充", "ai_focus": "待补充"}]

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
        typ = str(item.get("type", "")).strip()
        txt = _sanitize_single_line_field(str(item.get("text", "")), 450)
        focus = _sanitize_single_line_field(str(item.get("focus", "")), 520)
        if not txt:
            continue
        if _focus_is_weak(focus):
            focus = _build_professional_focus(txt)
        label = label_map.get(typ, "流程要求")
        process_rows.append({"requirement": f"{label}：{txt}", "focus": focus})
    if not process_rows:
        process_rows = [{"requirement": "待补充", "focus": "待补充"}]
    return criteria_rows, exclusion_rows, process_rows


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

    items_ai: Dict[str, Any] = {}
    if use_ai:
        basic_ai = _ai_extract_basic(text)
        if isinstance(basic_ai, dict) and not basic_ai.get("_error"):
            p = _sanitize_single_line_field(basic_ai.get("project_title", ""), 160)
            s = _sanitize_single_line_field(basic_ai.get("sponsor_name", ""), 100)
            c = _sanitize_single_line_field(basic_ai.get("protocol_code", ""), 50)
            v = _sanitize_single_line_field(basic_ai.get("version_no", ""), 30)
            d = _sanitize_single_line_field(basic_ai.get("version_date", ""), 40)
            ind = _sanitize_single_line_field(basic_ai.get("indication", ""), 100)
            ph = _sanitize_single_line_field(basic_ai.get("phase", ""), 30)
            sd = _sanitize_single_line_field(basic_ai.get("study_design", ""), 400)
            pe = _sanitize_single_line_field(basic_ai.get("primary_endpoint", ""), 400)
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
            items_ai = {}

    criteria_rows, exclusion_rows, process_rows = _rows_from_ai_or_rule(text, items_ai)

    ai_points = []
    for x in criteria_rows[:4]:
        ai_points.append(f"入选标准核查：{x['criterion']} → {x['ai_focus']}")
    for x in exclusion_rows[:4]:
        ai_points.append(f"排除标准核查：{x['criterion']} → {x['ai_focus']}")
    if primary_endpoint and primary_endpoint != "待补充主要终点":
        ai_points.append(f"主要终点核查：{primary_endpoint} → 查阅终点评估原始记录、评估时间窗、评分/判定依据和eCRF录入，确认源数据可溯源且与方案定义一致。")
    for x in process_rows[:4]:
        ai_points.append(f"流程要求核查：{x['requirement']} → {x['focus']}")
    ai_points += [
        "知情同意核查：重点关注伦理批准版本、签署时序、执行人员授权、受试者留存副本和重签触发条件。",
        "安全性核查：重点关注AE/SAE识别、分级、因果性判断、上报时效、随访闭环和医学处理记录。",
        "IMP/试验药物核查：重点关注接收、储存、温控、分发、使用、回收、销毁和账物一致性。",
        "数据溯源核查：重点关注HIS/LIS/PACS/EDC/eCRF之间的一致性、修改痕迹、签名日期和更正理由。",
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
