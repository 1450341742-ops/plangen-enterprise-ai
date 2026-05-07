from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List

CAPA_PATTERNS = {
    "培训": "培训类CAPA",
    "SOP": "SOP修订",
    "复核": "复核机制",
    "EDC": "数据核查",
    "药物": "IMP整改",
    "样本": "样本整改",
    "实验室": "实验室整改",
}


def extract_capa_patterns(text: str) -> Dict[str, List[dict]]:
    findings = []

    for keyword, category in CAPA_PATTERNS.items():
        count = len(re.findall(re.escape(keyword), text, flags=re.IGNORECASE))
        if count > 0:
            findings.append({
                "category": category,
                "keyword": keyword,
                "count": count,
                "recommendation": f"建议加强{category}相关管理与追踪。"
            })

    findings.sort(key=lambda x: x["count"], reverse=True)

    summary = Counter([f["category"] for f in findings])

    return {
        "patterns": findings,
        "summary": dict(summary)
    }
