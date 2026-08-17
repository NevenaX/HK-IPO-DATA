#!/usr/bin/env python
"""
apply_aliases.py — 将 investor_aliases.json 的规则重新应用到现有数据上。

用法:
  python scripts/apply_aliases.py

说明:
  读取 data/data.json，对每条记录的 investor_name 重新执行别名清洗，
  然后重新生成 data/data.js + data/data.json。
  不需要重新解析 PPT 文件，速度快。
"""

import json
import os
import re
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

ALIASES_PATH = os.path.join(BASE_DIR, "data", "investor_aliases.json")
DATA_JSON_PATH = os.path.join(BASE_DIR, "data", "data.json")
DATA_JS_PATH = os.path.join(BASE_DIR, "data", "data.js")

DATE_RE = re.compile(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$')


def load_aliases() -> dict:
    if os.path.exists(ALIASES_PATH):
        with open(ALIASES_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def normalize_investor_name(name: str, aliases: dict) -> str:
    """与 extract_pptx.py 完全一致的别名清洗逻辑。"""
    if not name:
        return name

    # 1. 精确匹配
    for old, new in aliases.get("exact", {}).items():
        if name == old:
            return new

    # 2. 包含关键字匹配
    for keyword, new in aliases.get("contains", {}).items():
        if keyword.lower() in name.lower():
            return new

    # 3. 子串替换
    for rule in aliases.get("substring", []):
        name = name.replace(rule["old"], rule["new"])
    name = name.strip()

    # 4. 不区分大小写匹配
    name_lower = name.lower()
    for old_lower, new in aliases.get("case_insensitive", {}).items():
        if name_lower == old_lower.lower():
            return new

    return name


def json_dumps(val) -> str:
    if val is None:
        return "null"
    s = str(val)
    s = s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "\\r").replace("\t", "\\t")
    return f'"{s}"'


def json_val(val) -> str:
    if val is None:
        return "null"
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return str(val)
    return str(val)


def records_to_js(records, source_files, date_cutoff=None):
    lines = [
        "// data/data.js — 自动生成，请勿手动修改",
        f"// 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    ]
    if source_files:
        lines.append(f"// 来源文件: {', '.join(source_files)}")
    if date_cutoff:
        lines.append(f"// 数据截止: {date_cutoff}")
    lines.append("")
    lines.append("window.cornerstoneData = [")

    for r in records:
        line = "  {"
        line += f'pricing_date: {json_dumps(r["pricing_date"])}, '
        line += f'stock_code: {json_dumps(r["stock_code"])}, '
        line += f'company_name: {json_dumps(r["company_name"])}, '
        line += f'industry: {json_dumps(r["industry"])}, '
        line += f'ipo_size_mn: {json_val(r["ipo_size_mn"])}, '
        line += f'cs_size_mn: {json_val(r["cs_size_mn"])}, '
        line += f'cs_ratio: {json_dumps(r["cs_ratio"])}, '
        line += f'investor_name: {json_dumps(r["investor_name"])}, '
        line += f'investor_category: {json_dumps(r["investor_category"])}, '
        line += f'amount_mn: {json_val(r["amount_mn"])}, '
        line += f'source_ppt: {json_dumps(r["source_ppt"])}, '
        line += f'slide_no: {r["slide_no"]}'
        line += "},"
        lines.append(line)

    lines.append("];")
    if date_cutoff:
        lines.append(f'window.dataCutoff = "{date_cutoff}";')
    lines.append("")
    return "\n".join(lines)


def main():
    print("=" * 60)
    print("  投资者别名清洗工具 — 重新应用 aliases 到现有数据")
    print("=" * 60)

    # 加载别名规则
    aliases = load_aliases()
    if not aliases:
        print("ERROR: data/investor_aliases.json 未找到或为空")
        sys.exit(1)

    exact_count = len(aliases.get("exact", {}))
    contains_count = len(aliases.get("contains", {}))
    substring_count = len(aliases.get("substring", []))
    ci_count = len(aliases.get("case_insensitive", {}))
    print(f"\n  别名规则: {exact_count} exact + {contains_count} contains + {substring_count} substring + {ci_count} case_insensitive")

    # 加载现有数据
    if not os.path.exists(DATA_JSON_PATH):
        print(f"ERROR: {DATA_JSON_PATH} 未找到")
        sys.exit(1)

    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        records = json.load(f)

    print(f"  原始数据: {len(records)} 条记录")
    print(f"  原始唯一投资者: {len(set(r['investor_name'] for r in records if r.get('investor_name')))}")

    # 应用别名清洗
    changes = 0
    for r in records:
        old_name = r.get("investor_name", "")
        new_name = normalize_investor_name(old_name, aliases)
        if new_name != old_name:
            changes += 1
            r["investor_name"] = new_name

    print(f"  名称变更: {changes} 条记录被清洗")

    # 获取来源文件和截止日期
    source_files = []
    seen = set()
    for r in records:
        f = r.get("source_ppt", "")
        if f and f not in seen:
            seen.add(f)
            source_files.append(f)

    dates = sorted(set(r.get("pricing_date", "") for r in records if r.get("pricing_date") and DATE_RE.match(r["pricing_date"])))
    date_cutoff = dates[-1] if dates else None

    # 写入 data.json
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    # 写入 data.js
    js_content = records_to_js(records, source_files, date_cutoff)
    with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)

    # 统计
    unique_investors = len(set(r["investor_name"] for r in records if r.get("investor_name")))
    total_amount = sum(r["amount_mn"] for r in records if r.get("amount_mn"))

    print(f"\n  [OK] 清洗完成!")
    print(f"  DONE -> {DATA_JS_PATH}")
    print(f"  DONE -> {DATA_JSON_PATH}")
    print(f"  总记录数: {len(records)}")
    print(f"  唯一项目数: {len(set(r['company_name'] for r in records if r.get('company_name')))}")
    print(f"  清洗后唯一投资者: {unique_investors}")
    print(f"  基石总额: ${total_amount:,.2f}M")
    print(f"  数据截止: {date_cutoff}")
    print()


if __name__ == "__main__":
    main()
