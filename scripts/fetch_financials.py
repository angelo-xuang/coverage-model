#!/usr/bin/env python3
"""
fetch_financials.py — 拉取个股近5年损益表数据
用法: python fetch_financials.py <ticker> [--proxy 127.0.0.1:7890] [--output json|csv]

数据源: yfinance + SEC EDGAR
自动将 FY 映射为自然年
"""

import sys
import json
import math
import argparse
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


def get_fy_to_calendar_offset(ticker_info: dict) -> int:
    """根据财年结束月份计算 FY→自然年的偏移量"""
    fy_end = ticker_info.get('financialCurrency', 'USD')
    # 大多数美股公司 FY12月结束，不需要偏移
    # 特殊情况：NVIDIA FY1月结束 → FY2026 = 自然年2025
    # yfinance 不直接返回 FY end month，用 fiscalYearEnd 判断
    return 0  # 默认不偏移，后面用实际日期判断


def map_fy_to_natural(date_str: str) -> int:
    """
    将财报日期映射为自然年。
    如果财报日期月份 <= 6月，该FY属于前一年；
    否则属于当年。
    例: '2025-01-26' (NVIDIA FY2026) → 2025
    """
    dt = datetime.strptime(date_str[:10], '%Y-%m-%d')
    if dt.month <= 6:
        return dt.year - 1
    return dt.year


def fetch_yfinance_income_stmt(ticker: str, proxy: str = None) -> dict:
    """通过 yfinance 拉取损益表"""
    import os
    if proxy:
        os.environ['HTTPS_PROXY'] = f'http://{proxy}'

    import yfinance as yf
    stock = yf.Ticker(ticker)

    # 获取年度损益表
    income_stmt = stock.financials  # 年度
    info = stock.info

    # 提取关键信息
    currency = info.get('financialCurrency', 'USD')
    fy_end_hint = info.get('fiscalYearEnd', None)

    # 将列名(日期字符串)映射到自然年
    result = {
        'ticker': ticker,
        'currency': currency,
        'fiscal_year_end': fy_end_hint,
        'data': {}
    }

    # 关键字段映射
    field_map = {
        'Total Revenue': 'revenue',
        'Cost Of Revenue': 'cogs',
        'Gross Profit': 'gross_profit',
        'Operating Revenue': 'operating_revenue',
        'Operating Expense': 'opex',
        'Operating Income': 'ebit',
        'Net Income': 'net_income',
        'Research And Development': 'rd',
        'Selling General And Administration': 'sga',
        'Interest Expense': 'interest_expense',
        'Tax Provision': 'tax',
        'Diluted EPS': 'diluted_eps',
        'Basic EPS': 'basic_eps',
    }

    for col in income_stmt.columns:
        nat_year = map_fy_to_natural(str(col.date()))
        year_label = f"{nat_year}A"

        row_data = {}
        for yf_name, our_name in field_map.items():
            if yf_name in income_stmt.index:
                val = income_stmt.loc[yf_name, col]
                if hasattr(val, 'item'):
                    val = val.item()
                elif isinstance(val, (int, float)):
                    val = float(val)
                else:
                    val = None
                # 过滤 NaN
                if isinstance(val, float) and math.isnan(val):
                    val = None
                row_data[our_name] = val

        result['data'][year_label] = row_data

    # 按年份排序
    result['data'] = dict(sorted(result['data'].items()))

    # 清除完全空的年份（所有字段都是 None）
    result['data'] = {
        k: v for k, v in result['data'].items()
        if any(val is not None for val in v.values())
    }

    # 拉取当前市值和PE
    result['current_market_cap'] = info.get('marketCap', None)
    result['current_pe'] = info.get('trailingPE', None)
    result['forward_pe'] = info.get('forwardPE', None)
    result['current_price'] = info.get('currentPrice', None)
    result['shares_outstanding'] = info.get('sharesOutstanding', None)
    result['fifty_two_week_high'] = info.get('fiftyTwoWeekHigh', None)
    result['fifty_two_week_low'] = info.get('fiftyTwoWeekLow', None)

    return result


def fetch_sec_edgar_revenue(ticker: str, proxy: str = None) -> dict:
    """
    从 SEC EDGAR 拉取收入数据作为补充/验证
    返回 {自然年: 收入} 字典
    """
    import os
    import requests

    proxies = {}
    if proxy:
        proxies = {'https': f'http://{proxy}', 'http': f'http://{proxy}'}

    # 先获取 CIK
    ticker_cik_url = "https://www.sec.gov/files/company_tickers.json"
    headers = {'User-Agent': 'coverage-model@example.com'}

    try:
        resp = requests.get(ticker_cik_url, headers=headers, proxies=proxies, timeout=30)
        resp.raise_for_status()
        tickers_data = resp.json()

        cik = None
        for entry in tickers_data.values():
            if entry.get('ticker', '').upper() == ticker.upper():
                cik = str(entry['cik_str']).zfill(10)
                break

        if not cik:
            return {}

        # 获取公司事实数据
        facts_url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
        resp = requests.get(facts_url, headers=headers, proxies=proxies, timeout=30)
        resp.raise_for_status()
        facts = resp.json()

        # 提取收入数据
        revenue_data = {}
        us_gaap = facts.get('facts', {}).get('us-gaap', {})

        # 尝试多个收入字段
        revenue_fields = ['Revenues', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'SalesRevenueNet']
        for field in revenue_fields:
            if field in us_gaap:
                units = us_gaap[field].get('units', {}).get('USD', [])
                for entry in units:
                    if entry.get('form') == '10-K' and entry.get('frame', '').startswith('CY'):
                        year = entry['frame'][2:6]
                        revenue_data[f"{year}A"] = entry['val']
                break

        return revenue_data
    except Exception as e:
        print(f"SEC EDGAR 获取失败: {e}", file=sys.stderr)
        return {}


def compute_margins(data: dict) -> dict:
    """计算利润率和增速"""
    years = sorted(data.keys())
    for i, year in enumerate(years):
        d = data[year]
        rev = d.get('revenue')
        gp = d.get('gross_profit')
        ebit = d.get('ebit')
        ni = d.get('net_income')

        if rev and rev != 0:
            d['gross_margin'] = round(gp / rev, 4) if gp else None
            d['ebit_margin'] = round(ebit / rev, 4) if ebit else None
            d['net_margin'] = round(ni / rev, 4) if ni else None

            if d.get('rd'):
                d['rd_ratio'] = round(d['rd'] / rev, 4)
            if d.get('sga'):
                d['sga_ratio'] = round(d['sga'] / rev, 4)

        # 计算增速
        if i > 0:
            prev_rev = data[years[i-1]].get('revenue')
            if prev_rev and prev_rev != 0 and rev:
                d['revenue_growth'] = round((rev - prev_rev) / prev_rev, 4)

    return data


def main():
    parser = argparse.ArgumentParser(description='拉取个股近5年财务数据')
    parser.add_argument('ticker', help='股票代码 (例: NVDA)')
    parser.add_argument('--proxy', default='127.0.0.1:7890', help='代理地址')
    parser.add_argument('--output', choices=['json', 'csv'], default='json', help='输出格式')
    parser.add_argument('--no-edgar', action='store_true', help='跳过SEC EDGAR')
    args = parser.parse_args()

    print(f"正在拉取 {args.ticker} 的财务数据...")

    # 1. yfinance 拉取
    print("  [1/2] yfinance 损益表...")
    result = fetch_yfinance_income_stmt(args.ticker, args.proxy)

    # 2. 计算利润率
    print("  [2/2] 计算利润率和增速...")
    result['data'] = compute_margins(result['data'])

    # 3. SEC EDGAR 补充（可选）
    if not args.no_edgar:
        print("  [补充] SEC EDGAR 收入验证...")
        edgar_rev = fetch_sec_edgar_revenue(args.ticker, args.proxy)
        if edgar_rev:
            result['edgar_revenue'] = edgar_rev

    # 4. 补充缺失年份：如果 yfinance 只有4年，用 EDGAR 填充第5年
    all_years = sorted(result['data'].keys())
    if len(all_years) == 4 and result.get('edgar_revenue'):
        first_year_int = int(all_years[0].replace('A', ''))
        missing_year = f"{first_year_int - 1}A"
        if missing_year in result['edgar_revenue']:
            result['data'][missing_year] = {'revenue': result['edgar_revenue'][missing_year]}
            print(f"  [补充] 用 EDGAR 填充了 {missing_year} 的收入数据", file=sys.stderr)

    # 只保留近5年
    all_years = sorted(result['data'].keys())
    if len(all_years) > 5:
        recent = all_years[-5:]
        result['data'] = {k: result['data'][k] for k in recent}

    # 输出
    if args.output == 'json':
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        # CSV 格式
        years = sorted(result['data'].keys())
        fields = ['revenue', 'revenue_growth', 'gross_profit', 'gross_margin',
                   'ebit', 'ebit_margin', 'net_income', 'net_margin', 'rd_ratio', 'sga_ratio']
        print(','.join(['指标'] + years))
        for f in fields:
            row = [f]
            for y in years:
                val = result['data'][y].get(f)
                row.append(f"{val:.4f}" if isinstance(val, float) else (str(val) if val else ''))
            print(','.join(row))

    # 输出摘要信息
    print(f"\n=== 摘要 ===", file=sys.stderr)
    print(f"股票: {args.ticker}", file=sys.stderr)
    print(f"数据年限: {', '.join(sorted(result['data'].keys()))}", file=sys.stderr)
    if result.get('current_market_cap'):
        print(f"当前市值: ${result['current_market_cap']/1e9:.1f}B", file=sys.stderr)
    if result.get('current_pe'):
        print(f"当前PE(TTM): {result['current_pe']:.1f}x", file=sys.stderr)
    if result.get('current_price'):
        print(f"当前股价: ${result['current_price']:.2f}", file=sys.stderr)


if __name__ == '__main__':
    main()
