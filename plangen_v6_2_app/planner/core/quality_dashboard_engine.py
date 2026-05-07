from __future__ import annotations

from collections import Counter
from typing import Dict, List

from planner.core.project_memory import load_project_memories
from planner.core.risk_profile_engine import build_risk_profile


def build_quality_dashboard() -> Dict[str, any]:
    memories = load_project_memories(limit=500)

    project_count = len(memories)

    risk_counter = Counter()

    for m in memories:
        text = "\n".join([
            m.get("source_text", ""),
            m.get("generated_markdown", ""),
        ])

        profile = build_risk_profile(text)

        for r in profile.get("top_risks", []):
            risk_counter[r["risk_category"]] += r["hit_count"]

    top_risks: List[dict] = []

    for name, count in risk_counter.most_common(10):
        top_risks.append({
            "risk_category": name,
            "count": count,
        })

    return {
        "project_count": project_count,
        "top_risks": top_risks,
    }
