#!/usr/bin/env python
"""
build_investor_search_aliases.py

把 data/investor_aliases.json 中用于数据清洗的 alias -> canonical 映射，
反向生成给网页搜索使用的 canonical -> [aliases...] 索引。

输出:
    data/investor_search_aliases.js

网页因此可以用中文名、英文名、简称或历史写法搜索同一个标准投资者。
"""

import json
from pathlib import Path
from collections import defaultdict

BASE = Path(__file__).resolve().parent.parent
INPUT = BASE / "data" / "investor_aliases.json"
OUTPUT = BASE / "data" / "investor_search_aliases.js"


def add(groups, canonical, alias):
    if not canonical or not alias:
        return
    canonical = str(canonical).strip()
    alias = str(alias).strip()
    if not canonical or not alias:
        return
    groups[canonical].add(canonical)
    groups[canonical].add(alias)


def main():
    cfg = json.loads(INPUT.read_text(encoding="utf-8"))
    groups = defaultdict(set)

    # 精确别名
    for alias, canonical in cfg.get("exact", {}).items():
        add(groups, canonical, alias)

    # contains 规则中的 keyword 也应能作为搜索词
    for keyword, canonical in cfg.get("contains", {}).items():
        add(groups, canonical, keyword)

    # 大小写不敏感的完整名称别名
    for alias, canonical in cfg.get("case_insensitive", {}).items():
        add(groups, canonical, alias)

    result = {
        canonical: sorted(values, key=lambda x: (x.lower() != canonical.lower(), x.lower()))
        for canonical, values in sorted(groups.items(), key=lambda kv: kv[0].lower())
    }

    js = (
        "// 自动生成，请勿手动修改；来源: data/investor_aliases.json\n"
        "window.investorSearchAliases = "
        + json.dumps(result, ensure_ascii=False, indent=2)
        + ";\n"
    )
    OUTPUT.write_text(js, encoding="utf-8")
    print(f"Generated {OUTPUT} ({len(result)} canonical investors)")


if __name__ == "__main__":
    main()
