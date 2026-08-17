#!/usr/bin/env python
"""
fetch_company_info.py — AKShare 批量抓取港股公司详情
=========================================================

对应 SKILL.md 的 Skill 5（公司上市信息详情侧栏）。

功能:
  1. 从 data/data.json 读取所有不重复的港股代码
  2. 用 AKShare 查询每家公司的上市信息与公司简介
  3. 从历史行情计算上市首日涨跌幅和 6 个月涨跌幅
  4. 与 PPT 已有数据（发行规模、基石规模等）合并
  5. 输出为 data/company_info.json

用法:
  python scripts/fetch_company_info.py          # 默认模式
  python scripts/fetch_company_info.py --force  # 强制重新抓取所有（跳过缓存）
  python scripts/fetch_company_info.py --quiet  # 静默模式

依赖:
  pip install akshare
"""

import os
import sys
import json
import time
import argparse
from datetime import datetime, timedelta

import akshare as ak
import pandas as pd


# =============================================================================
# 常量
# =============================================================================

DATA_DIR = os.path.dirname(os.path.dirname(__file__))
DATA_JSON_PATH = os.path.join(DATA_DIR, "data", "data.json")
COMPANY_INFO_PATH = os.path.join(DATA_DIR, "data", "company_info.json")
COMPANY_JS_PATH = os.path.join(DATA_DIR, "data", "company_info.js")

# 请求间隔（秒），避免触发反爬限制
REQUEST_INTERVAL = 0.5

# 涨跌幅计算中的6个月（交易日 ≈ 125 个交易日）
TRADING_DAYS_6M = 125


def to_hk_5digit(code: str) -> str:
    """
    将数据中的股票代码转换为 AKShare 5 位数字格式。

    例: "6675.HK" → "06675"
        "0068.HK" → "00068"
    """
    return code.replace(".HK", "").zfill(5)


def load_stock_codes() -> list[str]:
    """从 data.json 读取所有不重复的港股代码，按代码排序返回。"""
    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)
    codes = sorted(set(r["stock_code"] for r in raw if r.get("stock_code")))
    return codes


def fetch_security_profile(hk_code_5digit: str) -> dict:
    """
    从 AKShare 获取证券信息（上市日期、发行价、发行股数、上市板块等）。

    返回字段:
      - listing_date: 上市日期
      - issue_price: 发行价
      - issue_num: 发行股数
      - board: 上市板块（H股/红H股）
      - trade_unit: 每手股数
    """
    try:
        df = ak.stock_hk_security_profile_em(symbol=hk_code_5digit)
        if df is not None and len(df) > 0:
            row = df.iloc[0]
            return {
                "listing_date": str(row.iloc[2]) if pd.notna(row.iloc[2]) else None,
                "issue_price": float(row.iloc[4]) if pd.notna(row.iloc[4]) else None,
                "issue_num": float(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
                "board": str(row.iloc[3]) if pd.notna(row.iloc[3]) else None,
                "trade_unit": int(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
                "currency": str(row.iloc[7]) if pd.notna(row.iloc[7]) else None,
            }
    except Exception as e:
        pass
    return {}


def fetch_company_profile(hk_code_5digit: str) -> dict:
    """
    从 AKShare 获取公司简介。

    返回字段:
      - company_name_full: 公司全称（中文）
      - company_name_en: 英文名称
      - industry: 所属行业
      - company_profile: 公司简介文本
      - employees: 员工人数
      - chairman: 董事长
      - website: 公司网址
    """
    try:
        df = ak.stock_hk_company_profile_em(symbol=hk_code_5digit)
        if df is not None and len(df) > 0:
            row = df.iloc[0]
            # 列索引基于实际返回结构
            return {
                "company_name_full": str(row.iloc[0]) if pd.notna(row.iloc[0]) else None,
                "company_name_en": str(row.iloc[1]) if pd.notna(row.iloc[1]) else None,
                "industry": str(row.iloc[5]) if pd.notna(row.iloc[5]) else None,
                "chairman": str(row.iloc[6]) if pd.notna(row.iloc[6]) else None,
                "employees": str(row.iloc[8]) if pd.notna(row.iloc[8]) else None,
                "website": str(row.iloc[10]) if pd.notna(row.iloc[10]) else None,
                "company_profile": str(row.iloc[16]) if pd.notna(row.iloc[16]) else None,
            }
    except Exception as e:
        pass
    return {}


def fetch_price_data(hk_code_5digit: str, listing_date_str: str | None) -> dict:
    """
    从 AKShare 获取港股历史行情，计算上市首日和 6 个月涨跌幅。

    涨跌幅基准: (收盘价 - 发行价) / 发行价 × 100
    首日涨跌幅: 上市首日收盘价 vs 发行价
    6个月涨跌幅: 约 6 个月后的收盘价 vs 发行价

    返回字段:
      - first_day_close: 首日收盘价
      - first_day_change_pct: 首日涨跌幅（%）
      - last_close: 最新收盘价
      - latest_price_date: 最新价格日期
      - change_6m_pct: 6个月涨跌幅（%，若数据不足则为 None）
    """
    result = {
        "first_day_close": None,
        "first_day_change_pct": None,
        "last_close": None,
        "latest_price_date": None,
        "change_6m_pct": None,
    }

    if not listing_date_str:
        return result

    try:
        df = ak.stock_hk_daily(symbol=hk_code_5digit)
        if df is None or len(df) == 0:
            return result

        # 按日期排序
        df["date"] = pd.to_datetime(df["date"])
        df = df.sort_values("date").reset_index(drop=True)

        # 最新价格
        last_row = df.iloc[-1]
        result["last_close"] = float(last_row["close"]) if pd.notna(last_row["close"]) else None
        result["latest_price_date"] = str(last_row["date"].date()) if pd.notna(last_row["date"]) else None

        # 找到上市首日或之后的第一个交易日
        listing_date = pd.to_datetime(listing_date_str)
        first_day = df[df["date"] >= listing_date]
        if len(first_day) == 0:
            return result

        first_day_row = first_day.iloc[0]
        result["first_day_close"] = float(first_day_row["close"]) if pd.notna(first_day_row["close"]) else None

        # 查找6个月后的价格
        six_months_later = listing_date + timedelta(days=180)
        six_month_data = df[df["date"] >= six_months_later]
        if len(six_month_data) > 0:
            six_month_close = float(six_month_data.iloc[0]["close"]) if pd.notna(six_month_data.iloc[0]["close"]) else None
            if six_month_close is not None:
                result["six_month_close"] = six_month_close

    except Exception as e:
        pass

    return result


def calc_change_pct(close_price: float | None, issue_price: float | None) -> float | None:
    """计算涨跌幅百分比。"""
    if close_price is None or issue_price is None or issue_price == 0:
        return None
    return round((close_price - issue_price) / issue_price * 100, 2)


def get_ppt_data(codes: list[str]) -> dict:
    """
    从 data.json 提取 PPT 已有数据（发行规模、基石规模等），
    按股票代码聚合。
    """
    with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
        raw = json.load(f)

    company_data = {}
    for r in raw:
        code = r.get("stock_code")
        if not code:
            continue
        if code not in company_data:
            company_data[code] = {
                "stock_code": code,
                "company_name": r.get("company_name"),
                "industry_ppt": r.get("industry"),
                "total_ipo_size_mn": r.get("ipo_size_mn"),
                "cs_size_mn": r.get("cs_size_mn"),
                "cs_ratio": r.get("cs_ratio"),
                "pricing_date": r.get("pricing_date"),
            }

    return company_data


def main():
    parser = argparse.ArgumentParser(
        description="港股公司详情抓取脚本 — AKShare → company_info.json",
    )
    parser.add_argument("--force", "-f", action="store_true", help="强制重新抓取所有")
    parser.add_argument("--quiet", "-q", action="store_true", help="静默模式")
    args = parser.parse_args()

    # ---- 加载股票代码 ----
    codes = load_stock_codes()
    if not args.quiet:
        print(f"  [INFO] 共 {len(codes)} 只港股待查询")

    # ---- 加载已有数据（非 force 模式时，用于缓存跳过） ----
    existing = {}
    if os.path.exists(COMPANY_INFO_PATH) and not args.force:
        with open(COMPANY_INFO_PATH, "r", encoding="utf-8") as f:
            existing = json.load(f)

    # ---- 加载 PPT 基础数据 ----
    ppt_data = get_ppt_data(codes)

    # ---- 逐只查询 ----
    results = {}
    success_count = 0
    skip_count = 0
    refresh_count = 0
    fail_count = 0

    for i, code in enumerate(codes):
        hk_code = to_hk_5digit(code)
        base = ppt_data.get(code, {})

        # 缓存命中检查：有缓存的只刷新股价，不重新拉公司资料
        if code in existing:
            cached = existing[code]
            data_status = cached.get("data_status", {})
            if data_status.get("profile") == "ok" and data_status.get("issue_price") == "ok":
                if not args.quiet:
                    print(f"  [{i+1:2d}/{len(codes)}] {code} 刷新股价...", end=" ", flush=True)
                entry = dict(cached)
                listing_date_str = entry.get("listing_date")
                prices = fetch_price_data(hk_code, listing_date_str)
                issue_price = entry.get("issue_price")
                entry["last_close"] = prices.get("last_close")
                entry["latest_price_date"] = prices.get("latest_price_date")
                entry["first_day_close"] = prices.get("first_day_close")
                if prices.get("first_day_close") is not None and issue_price is not None:
                    entry["first_day_change_pct"] = calc_change_pct(prices["first_day_close"], issue_price)
                entry["six_month_close"] = prices.get("six_month_close")
                if prices.get("six_month_close") is not None and issue_price is not None:
                    entry["change_6m_pct"] = calc_change_pct(prices["six_month_close"], issue_price)
                else:
                    entry["change_6m_pct"] = None
                results[code] = entry
                refresh_count += 1
                if not args.quiet:
                    print(f"最新 {prices.get('latest_price_date', '?')}")
                continue

        if not args.quiet:
            print(f"  [{i+1:2d}/{len(codes)}] {code} ({hk_code}) 查询中...", end=" ", flush=True)

        data_status = {}

        # 1. 证券信息（上市日期、发行价、板块）
        sec = fetch_security_profile(hk_code)
        time.sleep(REQUEST_INTERVAL)

        # 2. 公司简介
        prof = fetch_company_profile(hk_code)
        time.sleep(REQUEST_INTERVAL)

        # 3. 历史行情（涨跌幅）
        listing_date_str = sec.get("listing_date")
        prices = fetch_price_data(hk_code, listing_date_str)
        time.sleep(REQUEST_INTERVAL)

        # 4. 计算涨跌幅
        issue_price = sec.get("issue_price")
        first_day_change = None
        change_6m = None

        if prices.get("first_day_close") is not None and issue_price is not None:
            first_day_change = calc_change_pct(prices["first_day_close"], issue_price)

        if prices.get("six_month_close") is not None and issue_price is not None:
            change_6m = calc_change_pct(prices["six_month_close"], issue_price)

        # 5. 汇总
        entry = {
            "stock_code": code,
            "company_name": base.get("company_name"),
            "company_name_en": prof.get("company_name_en"),
            "industry": prof.get("industry") or base.get("industry_ppt"),
            "board": sec.get("board"),
            "listing_date": sec.get("listing_date"),
            "issue_price": issue_price,
            "issue_num": sec.get("issue_num"),
            "trade_unit": sec.get("trade_unit"),
            "currency": sec.get("currency"),
            "total_ipo_size_mn": base.get("total_ipo_size_mn"),
            "cs_size_mn": base.get("cs_size_mn"),
            "cs_ratio": base.get("cs_ratio"),
            "first_day_close": prices.get("first_day_close"),
            "first_day_change_pct": first_day_change,
            "six_month_close": prices.get("six_month_close"),
            "change_6m_pct": change_6m,
            "last_close": prices.get("last_close"),
            "latest_price_date": prices.get("latest_price_date"),
            "chairman": prof.get("chairman"),
            "employees": prof.get("employees"),
            "website": prof.get("website"),
            "company_profile": prof.get("company_profile"),
            "data_status": {},
        }

        # 记录每个字段的获取状态
        entry["data_status"]["profile"] = "ok" if prof.get("company_profile") else "missing"
        entry["data_status"]["issue_price"] = "ok" if issue_price else "missing"
        entry["data_status"]["listing_date"] = "ok" if sec.get("listing_date") else "missing"
        entry["data_status"]["first_day_return"] = "ok" if first_day_change is not None else "missing"
        entry["data_status"]["six_month_return"] = "ok" if change_6m is not None else "missing"
        entry["data_status"]["sponsor"] = "missing"  # 暂无数据源
        entry["data_status"]["subscription_ratio"] = "missing"  # 暂无数据源

        results[code] = entry

        if prof.get("company_profile") or sec.get("listing_date"):
            success_count += 1
            if not args.quiet:
                print(f"OK")
        else:
            fail_count += 1
            if not args.quiet:
                print(f"⚠ 数据不完整")

    # ---- 输出 company_info.json ----
    os.makedirs(os.path.dirname(COMPANY_INFO_PATH) or ".", exist_ok=True)
    with open(COMPANY_INFO_PATH, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # ---- 输出 company_info.js（浏览器可直接加载） ----
    js_content = "// company_info.js — 自动生成，请勿手动修改\n"
    js_content += f"// 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
    js_content += f"// 覆盖 {len(results)} 只港股\n\n"
    js_content += "window.companyInfo = " + json.dumps(results, ensure_ascii=False, indent=2) + ";"
    with open(COMPANY_JS_PATH, "w", encoding="utf-8") as f:
        f.write(js_content)

    # ---- 统计 ----
    if not args.quiet:
        print()
        print(f"  DONE -> {COMPANY_INFO_PATH}")
        print(f"  Total: {len(codes)}, 首次查询: {success_count}, 股价刷新: {refresh_count}, 失败: {fail_count}")

        # 数据完整性统计
        fields_ok = {
            "profile": 0,
            "issue_price": 0,
            "listing_date": 0,
            "first_day_return": 0,
            "six_month_return": 0,
        }
        for r in results.values():
            ds = r.get("data_status", {})
            for k in fields_ok:
                if ds.get(k) == "ok":
                    fields_ok[k] += 1

        print()
        print(f"  ┌─────────────────────┬──────────┐")
        print(f"  │ 数据字段            │ 可用数   │")
        for k, v in fields_ok.items():
            label = {
                "profile": "公司简介",
                "issue_price": "发行价",
                "listing_date": "上市日期",
                "first_day_return": "首日涨跌幅",
                "six_month_return": "6个月涨跌幅",
            }.get(k, k)
            print(f"  │ {label:18s}    │ {v:5d}/{len(codes):<3d}  │")
        print(f"  └─────────────────────┴──────────┘")


if __name__ == "__main__":
    main()
