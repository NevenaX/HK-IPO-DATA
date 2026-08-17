#!/usr/bin/env python
"""
check_new_names.py — 检查新投资者名称是否被别名规则覆盖

功能：
  1. 接收新提取的记录列表，提取所有 investor_name
  2. 对照 investor_aliases.json，找出未被覆盖的名称
  3. 对未覆盖名称做模糊匹配，找出疑似同一机构的候选
  4. 输出可读报告 data/新增名称检查.md
  5. 不阻塞流程，仅提醒

用法（被 extract_pptx.py 调用）：
  from check_new_names import check_new_names
  check_new_names(new_records, aliases, quiet=False)

独立运行：
  python scripts/check_new_names.py <records.json>
"""

import json
import os
import sys
from datetime import datetime
from difflib import SequenceMatcher

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ALIASES_PATH = os.path.join(BASE_DIR, "data", "investor_aliases.json")
REPORT_PATH = os.path.join(BASE_DIR, "data", "新增名称检查.md")

# 已知无需处理的名称模式（个人名、英文名等）
KNOWN_CLEAN_PATTERNS = [
    "先生", "女士", "（个人）", "(个人)",
]


def load_aliases() -> dict:
    if os.path.exists(ALIASES_PATH):
        with open(ALIASES_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def normalize_investor_name(name: str, aliases: dict) -> str:
    """与 extract_pptx.py 一致的别名清洗逻辑。"""
    if not name:
        return name
    for old, new in aliases.get("exact", {}).items():
        if name == old:
            return new
    for keyword, new in aliases.get("contains", {}).items():
        if keyword.lower() in name.lower():
            return new
    for rule in aliases.get("substring", []):
        name = name.replace(rule["old"], rule["new"])
    name = name.strip()
    name_lower = name.lower()
    for old_lower, new in aliases.get("case_insensitive", {}).items():
        if name_lower == old_lower.lower():
            return new
    return name


def collect_canonical(aliases: dict) -> set:
    """收集所有规范化名称（别名规则的目标值）。"""
    names = set()
    for rule_type in ("exact", "contains", "case_insensitive"):
        for v in aliases.get(rule_type, {}).values():
            names.add(v)
    return names


def extract_zh(text: str) -> str:
    return "".join(c for c in text if '一' <= c <= '鿿')


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def find_similar(name: str, canonical_set: set) -> list[dict]:
    """找与 name 相似的规范名。"""
    name_zh = extract_zh(name)
    candidates = []

    for c_name in sorted(canonical_set):
        if c_name == name:
            continue
        c_zh = extract_zh(c_name)
        best_sim = 0.0
        match_type = ""

        # 中文相同 → 高度疑似
        if name_zh and c_zh and name_zh == c_zh:
            best_sim = 0.92
            match_type = "中文名相同"
        # 中文包含
        elif name_zh and c_zh and (name_zh in c_zh or c_zh in name_zh):
            sim = min(len(name_zh), len(c_zh)) / max(len(name_zh), len(c_zh))
            if sim > 0.4:
                best_sim = 0.75 + sim * 0.1
                match_type = "中文名包含"

        # 全名模糊匹配
        full_sim = similarity(name, c_name)
        if full_sim > 0.5 and full_sim > best_sim:
            best_sim = full_sim
            match_type = f"模糊匹配({full_sim:.0%})"

        if best_sim >= 0.5:
            candidates.append({
                "name": c_name,
                "similarity": round(best_sim, 3),
                "match_type": match_type,
            })

    candidates.sort(key=lambda x: x["similarity"], reverse=True)
    return candidates[:3]


def is_personal_name(name: str) -> bool:
    """判断是否为人名（通常不需要别名规则）。"""
    for p in KNOWN_CLEAN_PATTERNS:
        if p in name:
            return True
    # 纯英文名 + 括号内个人名
    if "（" in name or "(" in name:
        return False  # 让具体逻辑判断
    # 单人汉字名（2-4字纯中文）
    zh = extract_zh(name)
    if zh and len(zh) <= 4 and len(zh) == len(name.strip()):
        return True
    return False


def check_new_names(
    new_records: list[dict],
    aliases: dict = None,
    quiet: bool = False
) -> dict:
    """
    检查新记录中的投资者名称。

    Args:
        new_records: 新提取的记录列表
        aliases: 别名规则（不传则自动加载）
        quiet: 静默模式

    Returns:
        {"has_uncovered": bool, "total_new": int, "suggestions": [...]}
    """
    if aliases is None:
        aliases = load_aliases()
    if not aliases:
        if not quiet:
            print("  [WARN] investor_aliases.json 未找到，跳过名称检查")
        return {"has_uncovered": False, "total_new": 0, "suggestions": [], "uncovered_details": []}

    canonical = collect_canonical(aliases)

    # 收集新记录中所有 unique investor_name
    new_names = set()
    name_counts = {}
    for r in new_records:
        name = r.get("investor_name", "")
        if name:
            new_names.add(name)
            name_counts[name] = name_counts.get(name, 0) + 1

    if not new_names:
        return {"has_uncovered": False, "total_new": 0, "suggestions": [], "uncovered_details": []}

    # 逐个检查
    uncovered = []
    for name in sorted(new_names):
        normalized = normalize_investor_name(name, aliases)

        # 已覆盖（别名规则改变了名称）
        if normalized != name:
            continue

        # 本身就是规范名
        if name in canonical:
            continue

        # 个人名跳过
        if is_personal_name(name):
            continue

        # 未覆盖 → 找候选
        candidates = find_similar(name, canonical)
        uncovered.append({
            "name": name,
            "count": name_counts[name],
            "candidates": candidates,
            "best_match": candidates[0]["name"] if candidates else None,
        })

    result = {
        "has_uncovered": len(uncovered) > 0,
        "total_new": len(new_names),
        "suggestions": len(uncovered),
        "uncovered_details": uncovered,
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    if not quiet:
        if result["has_uncovered"]:
            print(f"  [NAME_CHECK] 发现 {len(uncovered)} 个新名称可能需添加别名规则:")
            for item in uncovered:
                if item["candidates"]:
                    print(f"      {item['name']} ({item['count']}次) → 疑似: {item['candidates'][0]['name']}")
                else:
                    print(f"      {item['name']} ({item['count']}次) → 未匹配到候选")
            print(f"  [NAME_CHECK] 详情见 data/新增名称检查.md")
        else:
            print(f"  [NAME_CHECK] 所有 {result['total_new']} 个新名称均已覆盖")

    # 有发现则保存报告
    if result["has_uncovered"]:
        _save_report(result)

    return result


def _save_report(result: dict):
    """保存可读报告。"""
    lines = []
    lines.append("# 新增投资者名称检查报告")
    lines.append("")
    lines.append(f"检查时间: {result['checked_at']}")
    lines.append(f"新增记录检查数: —")
    lines.append(f"新增唯一投资者数: {result['total_new']}")
    lines.append(f"建议关注: {result['suggestions']} 个")
    lines.append("")

    if not result["has_uncovered"]:
        lines.append("所有名称均已覆盖，无需处理。")
        lines.append("")
        with open(REPORT_PATH, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return

    lines.append("## 建议关注的新名称")
    lines.append("")
    lines.append("以下名称未被现有别名规则覆盖，请确认是否需要新增规则：")
    lines.append("")
    for item in result["uncovered_details"]:
        lines.append(f"- **{item['name']}**（出现 {item['count']} 次）")
        if item["candidates"]:
            lines.append(f"  - 疑似同一机构：")
            for c in item["candidates"]:
                flag = " ⬅ 建议" if c == item["candidates"][0] else ""
                lines.append(f"    - `{c['name']}`（{c['match_type']}，相似度 {c['similarity']:.0%}）{flag}")
            lines.append(f'  - 如需合并，在 investor_aliases.json 的 "exact" 中添加：')
            lines.append(f'    "{item["name"]}": "{item["best_match"]}"')
        else:
            lines.append(f"  - 未发现明显相似项，如确认是新机构则无需处理")
        lines.append("")

    lines.append("---")
    lines.append("*本报告由 extract_pptx.py 自动生成，仅供参考，不阻塞流程。*")
    lines.append("")

    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


# =============================================================================
# 独立运行入口
# =============================================================================

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python scripts/check_new_names.py <records.json>")
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        records = json.load(f)

    result = check_new_names(records, quiet=False)
    sys.exit(1 if result["has_uncovered"] else 0)
