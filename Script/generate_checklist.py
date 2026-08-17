#!/usr/bin/env python
"""
generate_checklist.py — 更新后自动生成名称清洗确认清单（给同事看）
===================================================================

从 data/新增名称检查.md 中读取原始报告，生成同事友好的 .txt 版本，
按"明显可合并 / 需要确认 / 新机构"分类，同事可以直接双击打开。

被 更新数据.bat 自动调用，不阻塞流程。
"""

import os
import re
import sys
import json
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPORT_MD = os.path.join(BASE_DIR, "data", "新增名称检查.md")
CHECKLIST_TXT = os.path.join(BASE_DIR, "名称清洗清单.txt")
ALIASES_PATH = os.path.join(BASE_DIR, "data", "investor_aliases.json")


def load_aliases() -> dict:
    """加载别名规则，用于判断哪些已经是标准名。"""
    if os.path.exists(ALIASES_PATH):
        try:
            with open(ALIASES_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def collect_canonical_names(aliases: dict) -> set:
    """收集所有规范化名称。"""
    names = set()
    for rule_type in ("exact", "contains", "case_insensitive"):
        for v in aliases.get(rule_type, {}).values():
            names.add(v)
    return names


def extract_zh(text: str) -> str:
    return "".join(c for c in text if '一' <= c <= '鿿')


def has_zh(text: str) -> bool:
    return bool(extract_zh(text))


def parse_raw_report(filepath: str) -> list[dict]:
    """
    解析 新增名称检查.md 中的条目。
    返回: [{"name": str, "count": int, "candidates": [{"name": str, "sim": str, "type": str}], "best_match": str}]
    """
    if not os.path.exists(filepath):
        return []

    with open(filepath, "r", encoding="utf-8") as f:
        content = f.read()

    items = []
    # 匹配条目: - **名称**（出现 N 次）
    pattern = re.compile(
        r'- \*\*(.+?)\*\*（出现 (\d+) 次）\n'
        r'(?:  - 疑似同一机构：\n'
        r'(.*?)'
        r'  - 如需合并.*?\n'
        r'    "(.*?)": "(.*?)"\n)?'
        r'(?:  - 未发现明显相似项.*?\n)?',
        re.DOTALL
    )

    # 更简单的逐行解析
    lines = content.split("\n")
    current_item = None

    for i, line in enumerate(lines):
        # 匹配条目行: - **名称**（出现 N 次）
        m = re.match(r'- \*\*(.+?)\*\*（出现 (\d+) 次）', line)
        if m:
            if current_item:
                items.append(current_item)
            current_item = {
                "name": m.group(1),
                "count": int(m.group(2)),
                "candidates": [],
                "best_match": None,
                "no_match": False,
            }
            continue

        if current_item is None:
            continue

        # 匹配候选行: - `候选名`（类型，相似度 X%）
        m = re.match(r'\s*-\s*`(.+?)`（(.+?)，相似度 (\d+)%）', line)
        if m:
            current_item["candidates"].append({
                "name": m.group(1),
                "match_type": m.group(2),
                "similarity": int(m.group(3)),
            })
            if current_item["best_match"] is None:
                current_item["best_match"] = m.group(1)
            continue

        # 匹配"未发现明显相似项"
        if "未发现明显相似项" in line:
            current_item["no_match"] = True

    if current_item:
        items.append(current_item)

    return items


def classify_item(item: dict, canonical: set) -> str:
    """
    对条目分类:
      - category_1: 明显可合并（中文名相同，高度疑似）
      - category_2: 需要人工判断
      - category_3: 新机构（无需处理）
    """
    name = item["name"]
    name_zh = extract_zh(name)

    # 没有候选 → 新机构
    if item["no_match"] or not item["candidates"]:
        return "category_3"

    # 最好的候选
    best = item["candidates"][0]
    best_name = best["name"]
    best_sim = best["similarity"]
    best_type = best["match_type"]

    # 中文名完全相同 → 明显是同一机构
    if best_type == "中文名相同" and best_sim >= 90:
        return "category_1"

    # 中文名包含，且相似度>75 → 大概率正确
    if "中文名包含" in best_type and best_sim >= 75:
        return "category_1"

    # 模糊匹配且sim >= 80
    if "模糊匹配" in best_type and best_sim >= 80:
        return "category_1"

    # 有中文名匹配 → 需要判断
    if name_zh and has_zh(best_name):
        return "category_2"

    # 纯英文模糊匹配，相似度不高 → 需要判断
    if "模糊匹配" in best_type and best_sim >= 60:
        return "category_2"

    # 其他 → 新机构
    return "category_3"


def generate_txt(items: list[dict], canonical: set) -> str:
    """生成同事友好的 .txt 确认清单。"""
    lines = []
    lines.append("=" * 60)
    lines.append("  港股 IPO 基石投资者数据库 — 名称清洗确认清单")
    lines.append(f"  生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 60)
    lines.append("")
    lines.append("【操作说明】")
    lines.append("  1. 看下面清单，找到你觉得应该是同一机构的不同写法")
    lines.append("  2. 把确认结果微信发给我，格式如：")
    lines.append('     "第1条合并" 或 "第5条保留原名"')
    lines.append("  3. 我收到后会更新规则，下次更新时自动生效")
    lines.append("")

    # 分类
    cat1, cat2, cat3 = [], [], []
    for item in items:
        cat = classify_item(item, canonical)
        if cat == "category_1":
            cat1.append(item)
        elif cat == "category_2":
            cat2.append(item)
        else:
            cat3.append(item)

    # ---- 第一组：明显可合并 ----
    if cat1:
        lines.append("-" * 60)
        lines.append("【第一组：大概率是同一机构，可直接合并】")
        lines.append("")
        for i, item in enumerate(cat1, 1):
            best = item["candidates"][0]
            lines.append(f"  {i}. {item['name']}（{item['count']}次）")
            lines.append(f"     → {best['name']}（{best['match_type']} {best['similarity']}%）")
            lines.append(f"     确认合并? ___  保留原名? ___")
            lines.append("")

    # ---- 第二组：需要确认 ----
    if cat2:
        lines.append("-" * 60)
        lines.append("【第二组：需要你判断的，不确定的问我】")
        lines.append("")
        for i, item in enumerate(cat2, len(cat1) + 1):
            lines.append(f"  {i}. {item['name']}（{item['count']}次）")
            for c in item["candidates"]:
                flag = "  ← 建议" if c == item["candidates"][0] else ""
                lines.append(f"     → 疑似: {c['name']}（{c['match_type']} {c['similarity']}%）{flag}")
            lines.append(f"     确认合并? ___  保留原名? ___")
            lines.append("")

    # ---- 第三组：新机构 ----
    if cat3:
        lines.append("-" * 60)
        lines.append(f"【第三组：以下 {len(cat3)} 个为新机构，无需处理（仅供参考）】")
        lines.append("")
        for i, item in enumerate(cat3, len(cat1) + len(cat2) + 1):
            lines.append(f"  {i}. {item['name']}（{item['count']}次）")

    lines.append("")
    lines.append("-" * 60)
    lines.append("")

    # 汇总统计
    total = len(items)
    lines.append(f"【汇总】共 {total} 个未覆盖名称")
    if cat1:
        lines.append(f"  - 建议合并: {len(cat1)} 个（第一组）")
    if cat2:
        lines.append(f"  - 待确认: {len(cat2)} 个（第二组）")
    if cat3:
        lines.append(f"  - 新机构: {len(cat3)} 个（无需处理）")
    lines.append("")
    lines.append("确认后把标记结果截图发给我即可。")
    lines.append("")

    return "\n".join(lines)


def print_summary(items: list[dict], canonical: set):
    """在黑窗口打印简要摘要。"""
    cat1, cat2, cat3 = [], [], []
    for item in items:
        cat = classify_item(item, canonical)
        if cat == "category_1":
            cat1.append(item)
        elif cat == "category_2":
            cat2.append(item)
        else:
            cat3.append(item)

    print()
    print("  ===============================================")
    print("  名称清洗确认清单已生成 -> 名称清洗清单.txt")
    print("  ===============================================")
    print(f"  共 {len(items)} 个未覆盖的投资者名称：")
    if cat1:
        print(f"    建议合并: {len(cat1)} 个（清单第一组）")
    if cat2:
        print(f"    待您确认: {len(cat2)} 个（清单第二组）")
    if cat3:
        print(f"    新机构:   {len(cat3)} 个（无需处理，仅供参考）")
    print()
    print("  操作：")
    if cat1 or cat2:
        print("    1. 打开 名称清洗清单.txt")
        print('    2. 在每条后面写"合并"或"保留"')
        print("    3. 把文件发给我就行")
    else:
        print("    本次无新名称需要处理，一切正常！")
    print("  ===============================================")
    print()


def main():
    # 检查是否有原始报告
    if not os.path.exists(REPORT_MD):
        print("  [CHECKLIST] 无新增名称检查报告，跳过清单生成")
        return

    # 读取原始报告
    with open(REPORT_MD, "r", encoding="utf-8") as f:
        content = f.read()

    # 检查是否为空报告（所有名称已覆盖）
    if "所有名称均已覆盖" in content or "建议关注: 0 个" in content:
        # 删除旧清单（如果有）
        if os.path.exists(CHECKLIST_TXT):
            os.remove(CHECKLIST_TXT)
        print("  [CHECKLIST] 所有名称已覆盖，无需确认清单")
        return

    # 解析条目
    items = parse_raw_report(REPORT_MD)
    if not items:
        print("  [CHECKLIST] 报告解析异常，跳过清单生成")
        return

    # 收集标准名
    aliases = load_aliases()
    canonical = collect_canonical_names(aliases)

    # 生成清单
    txt_content = generate_txt(items, canonical)
    with open(CHECKLIST_TXT, "w", encoding="utf-8") as f:
        f.write(txt_content)

    # 打印摘要
    print_summary(items, canonical)


if __name__ == "__main__":
    main()
