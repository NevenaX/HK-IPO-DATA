#!/usr/bin/env python
"""
bundle_standalone.py — 将数据嵌入 HTML，生成单文件版本

从 data/data.js 和 data/company_info.js 读取数据，
嵌入到 hk_cornerstone_investors.html 的 <script> 标签中，
输出 hk_cornerstone_investors_standalone.html（单文件，即开即用）。

用法:
  python scripts/bundle_standalone.py
"""

import os
import re
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SRC_HTML = os.path.join(BASE_DIR, "hk_cornerstone_investors.html")
DATA_JS = os.path.join(BASE_DIR, "data", "data.js")
COMPANY_JS = os.path.join(BASE_DIR, "data", "company_info.js")
IPO_DETAILS_JS = os.path.join(BASE_DIR, "data", "ipo_details.js")
OUT_HTML = os.path.join(BASE_DIR, "hk_cornerstone_investors_standalone.html")


def read_file(path, label):
    if not os.path.exists(path):
        print(f"  [ERROR] 未找到 {label}: {path}")
        sys.exit(1)
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def main():
    print(f"  生成单文件 HTML ...")

    html = read_file(SRC_HTML, "源 HTML")
    data_js = read_file(DATA_JS, "data.js")
    company_js = read_file(COMPANY_JS, "company_info.js")
    ipo_js = read_file(IPO_DETAILS_JS, "ipo_details.js")

    # 1. 内联 data.js（替换 <script src="data/data.js">）
    html = html.replace(
        '<script src="data/data.js"></script>',
        f"<script>\n{data_js}\n</script>"
    )

    # 2. 内联 company_info.js（替换 <script src="data/company_info.js">）
    html = html.replace(
        '<script src="data/company_info.js"></script>',
        f"<script>\n{company_js}\n</script>"
    )

    # 3. 内联 ipo_details.js（替换 <script src="data/ipo_details.js">）
    html = html.replace(
        '<script src="data/ipo_details.js"></script>',
        f"<script>\n{ipo_js}\n</script>"
    )

    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    size_kb = os.path.getsize(OUT_HTML) / 1024
    print(f"  DONE -> {OUT_HTML} ({size_kb:.0f} KB)")
    print(f"  将此文件直接发给领导，双击即可打开")


if __name__ == "__main__":
    main()
