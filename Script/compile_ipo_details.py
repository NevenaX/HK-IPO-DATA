#!/usr/bin/env python
"""
compile_ipo_details.py — 自动编译港股IPO补充详情数据
====================================================
自动从 data/data.json 发现所有股票代码，
与 data/ipo_details.json（已有数据）合并，
新公司自动追加空条目，无需手动修改代码。

数据由 Claude 通过东方财富 MCP（mx_hk_finance_data）查询后，
直接更新到 data/ipo_details.json，再运行此脚本重新生成 JS。

用法:
  python scripts/compile_ipo_details.py
"""

import json
import os
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
DATA_JSON_PATH = os.path.join(DATA_DIR, "data.json")
IPO_DETAILS_PATH = os.path.join(DATA_DIR, "ipo_details.json")
IPO_DETAILS_JS_PATH = os.path.join(DATA_DIR, "ipo_details.js")

# 新公司空条目模板（字段留空，不阻塞流程）
EMPTY_TEMPLATE = {
    "issue_price": None,
    "net_proceeds_mn": None,
    "sponsors": None,
    "stabilizing_agent": None,
    "lot_size": None,
    "board": None,
    "overallotment_shares": None,
    "public_sub_multiple": None,
    "intl_sub_multiple": None,
    "incorporation_date": None,
    "registered_capital": None,
}


def main():
    # 1. 从 data.json 读取所有股票代码
    if not os.path.exists(DATA_JSON_PATH):
        print(f"  [ERROR] 未找到 {DATA_JSON_PATH}，请先运行 extract_pptx.py")
        return

    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        records = json.load(f)

    all_codes = set()
    for r in records:
        code = r.get("stock_code", "")
        if code:
            all_codes.add(code)

    if not all_codes:
        print(f"  [ERROR] {DATA_JSON_PATH} 中未找到任何股票代码")
        return

    # 2. 加载现有 IPO 详情
    existing = {}
    if os.path.exists(IPO_DETAILS_PATH):
        with open(IPO_DETAILS_PATH, "r", encoding="utf-8") as f:
            existing = json.load(f)

    # 3. 合并：保留旧数据，新代码自动追加空条目
    merged = {}
    for code in existing:
        if code in all_codes:
            merged[code] = existing[code]

    new_count = 0
    for code in sorted(all_codes):
        if code not in merged:
            merged[code] = dict(EMPTY_TEMPLATE)
            new_count += 1

    sorted_merged = dict(sorted(merged.items()))

    # 4. 输出 JSON
    with open(IPO_DETAILS_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted_merged, f, ensure_ascii=False, indent=2)

    # 5. 输出 JS（浏览器加载用）
    output_js(sorted_merged)

    # 6. 统计
    missing = sorted(c for c, v in sorted_merged.items() if v.get("sponsors") is None)
    filled = len(sorted_merged) - len(missing)

    print(f"  DONE -> {IPO_DETAILS_JS_PATH}")
    print(f"  DONE -> {IPO_DETAILS_PATH}")
    print(f"  共 {len(sorted_merged)} 只港股")
    print(f"  已有 IPO 详情: {filled} 只")
    if new_count > 0:
        print(f"  [NEW] 自动追加新公司: {new_count} 只（字段为空，待 MCP 补充）")
    if missing:
        new_missing = [c for c in missing if c not in existing]
        old_missing = [c for c in missing if c in existing]
        if new_missing:
            print(f"  [MCP] 新公司待补充: {len(new_missing)} 只")
            for c in new_missing:
                print(f"    {c}")
        if old_missing:
            print(f"  [MCP] 历史待补充: {len(old_missing)} 只")
    else:
        print(f"  [OK] 全部 {len(sorted_merged)} 只港股均有 IPO 详情数据")


def output_js(data):
    """生成浏览器可加载的 JS 文件"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        f"// ipo_details.js — 自动生成，请勿手动修改",
        f"// 生成时间: {timestamp}",
        f"// 覆盖 {len(data)} 只港股",
        f"",
        f"window.ipoDetails = {json.dumps(data, ensure_ascii=False, indent=2)};",
        f"",
    ]
    with open(IPO_DETAILS_JS_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


if __name__ == "__main__":
    main()
