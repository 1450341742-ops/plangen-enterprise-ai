from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List

HIGH_RISK_TERMS = {
    "SAE": "安全性",
    "死亡": "安全性",
    "随机": "随机化",
    "入选": "入排",
    "排除": "入排",
    "方案偏离": "方案执行",
    "药物": "IMP",
    "样本": "生物样本",
    "实验室": "中心实验室",
    "EDC": "数据一致性",
    "ICF": "知情同意",
    "时间窗": "时间窗",
}


def build_risk_profile(text: str) -> Dict[str, any]:
    categories: List[str] = []

    for term, cat in HIGH_RISK_TERMS.items():
        count = len(re.findall(re.escape(term), text, flags=re.IGNORECASE))
        categories.extend([cat] * count)

    counter = Counter(categories)

    top_risks = []
    for name, count in counter.most_common(8):
        top_risks.append({
            "risk_category": name,
            "hit_count": count,
            "risk_level": "高" if count >= 5 else "中" if count >= 2 else "低"
        })

    return {
        "top_risks": top_risks,
        "risk_summary": "；".join([
            f"{r['risk_category']}({r['risk_level']})"
            for r in top_risks
        ])
    }
