#!/usr/bin/env python3
"""
fetch_financials.py — 拉取个股近5年损益表数据
用法: python fetch_financials.py <ticker> [--proxy 127.0.0.1:7890] [--output json|csv]

数据源: yfinance + SEC EDGAR
自动将 FY 映射为自然年（根据公司实际财年定义）
支持多币种自动转换和语言切换
"""

import sys
import json
import math
import argparse
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

from company_config import detect_region, get_fx_rate, choose_unit


def analyze_fiscal_year(ticker_info: dict, income_stmt_columns: list) -> dict:
    """
    分析公司的财年定义和周期，生成精确的FY→自然年映射规则。

    不能简单地把"FY减一"当年自然年。不同公司的披露周期不同：
    - NVIDIA: FY结束于1月底，FY2026 = 自然年2025
    - Apple: FY结束于9月底，FY2024 = 自然年2024的10月-2024年9月
    - Microsoft: FY结束于6月底，FY2025 = 自然年2024的7月-2025年6月
    - 大多数公司: FY结束于12月，与自然年一致

    Returns:
    {
        'fy_end_month': int,           # 财年结束月份
        'fy_end_day': int,             # 财年结束日（如果知道）
        'fiscal_year_end': str,        # 如 "01-31", "06-30", "09-28", "12-31"
        'mapping_rule': str,           # 映射规则说明
        'is_calendar_year': bool,      # 是否与自然年一致
        'natural_year_offset': int,    # 自然年偏移（0=当年，-1=上一年）
        'fiscal_year_note': str,       # 完整说明
        'date_mappings': {date_str: natural_year},  # 每个具体日期的映射
    }
    """
    result = {
        'fy_end_month': 12,
        'fy_end_day': 31,
        'fiscal_year_end': '12-31',
        'mapping_rule': 'FY与自然年一致',
        'is_calendar_year': True,
        'natural_year_offset': 0,
        'fiscal_year_note': '',
        'date_mappings': {},
    }

    # Step 1: 从 yfinance info 获取 FY 结束信息
    fy_end_str = ticker_info.get('fiscalYearEnd', '')
    if fy_end_str:
        try:
            parts = fy_end_str.split('-')
            fy_month = int(parts[0])
            fy_day = int(parts[1])
            result['fy_end_month'] = fy_month
            result['fy_end_day'] = fy_day
            result['fiscal_year_end'] = f"{fy_month:02d}-{fy_day:02d}"
        except (ValueError, IndexError):
            pass

    # Step 2: 从实际财报日期推断
    # yfinance 的 income_stmt 列是实际的报告期结束日期
    report_months = []
    for col in income_stmt_columns:
        date_str = str(col.date())
        dt = datetime.strptime(date_str[:10], '%Y-%m-%d')
        report_months.append(dt.month)

    if report_months:
        # 看报告期结束月份的众数
        from collections import Counter
        month_counts = Counter(report_months)
        most_common_month = month_counts.most_common(1)[0][0]

        # 如果 yfinance 没给 FY end，用实际日期推断
        if not fy_end_str:
            result['fy_end_month'] = most_common_month
            result['fiscal_year_end'] = f"{most_common_month:02d}-??"

    fy_month = result['fy_end_month']

    # Step 3: 确定映射规则
    if fy_month == 12:
        result['is_calendar_year'] = True
        result['natural_year_offset'] = 0
        result['mapping_rule'] = 'FY结束于12月，与自然年一致，直接使用FY标签年份'
    elif fy_month <= 6:
        # FY结束于1-6月：该FY的大部分经营活动发生在上一个自然年
        # 例如 NVIDIA FY end Jan → FY2026 = CY2025
        # Microsoft FY end Jun → FY2025 = CY2024 (但实际是 2024.7-2025.6)
        result['is_calendar_year'] = False
        result['natural_year_offset'] = -1
        result['mapping_rule'] = (
            f'FY结束于{fy_month}月，该FY的大部分经营活动属于上一个自然年。'
            f'映射规则：自然年 = FY标签年 - 1'
        )
    else:
        # FY结束于7-11月
        result['is_calendar_year'] = False
        result['natural_year_offset'] = 0
        result['mapping_rule'] = (
            f'FY结束于{fy_month}月，该FY的经营活动跨越两个自然年。'
            f'映射规则：自然年 ≈ FY标签年（以大部分经营活动所在年份为准）'
        )

    # Step 4: 对每个具体日期做精确映射
    for col in income_stmt_columns:
        date_str = str(col.date())
        dt = datetime.strptime(date_str[:10], '%Y-%m-%d')
        nat_year = _map_single_date(dt, fy_month)
        result['date_mappings'][date_str[:10]] = nat_year

    # Step 5: 生成完整说明
    result['fiscal_year_note'] = _generate_fy_note(
        result['fy_end_month'],
        result['fy_end_day'],
        result['mapping_rule'],
        result['date_mappings'],
        ticker_info.get('longName', ticker_info.get('shortName', ''))
    )

    return result


def _map_single_date(report_date: datetime, fy_end_month: int) -> int:
    """
    将单个财报日期映射为自然年。

    核心逻辑：
    - FY结束于12月：直接用报告日期的年份
    - FY结束于1-6月：该FY属于前一个自然年
    - FY结束于7-11月：该FY属于当年（大部分经营活动在当年）

    但不是简单地看月份，而是要理解：
    财报日期 = 该FY结束的那一天
    所以 FY 的经营活动期间 = (报告日期 - 12个月, 报告日期]
    """
    if fy_end_month == 12:
        return report_date.year
    elif fy_end_month <= 6:
        # FY 结束于 1-6 月，该 FY 的经营活动主要在前一年
        # 例: 报告日期 2025-01-26 → 自然年 2024（NVIDIA FY2025）
        # 但这里需要注意：yfinance 返回的列标签已经是实际日期
        # FY2026 (ending Jan 2026) 的收入主要在 2025年产生
        return report_date.year - 1
    else:
        # FY 结束于 7-11 月
        # 例: Apple FY2024 结束于 Sep 2024 → 自然年 2024
        return report_date.year


def _generate_fy_note(fy_month: int, fy_day: int, rule: str,
                       mappings: dict, company_name: str) -> str:
    """生成给用户看的财年说明"""
    lines = [
        f"【财年定义】",
        f"公司: {company_name}",
        f"财年结束: 每年{fy_month}月",
        f"映射规则: {rule}",
        f"",
        f"【具体映射】",
    ]
    for date_str in sorted(mappings.keys()):
        lines.append(f"  财报日期 {date_str} → 自然年 {mappings[date_str]}")

    lines.append("")
    lines.append("注意：以上映射基于财报日期自动判断，如公司有特殊披露周期请手动核实。")

    return '\n'.join(lines)


def fetch_yfinance_income_stmt(ticker: str, proxy: str = None) -> dict:
    """通过 yfinance 拉取损益表"""
    import os
    if proxy:
        os.environ['HTTPS_PROXY'] = f'http://{proxy}'

    import yfinance as yf
    stock = yf.Ticker(ticker)

    income_stmt = stock.financials
    info = stock.info

    # === 检测公司地区和货币配置 ===
    company_cfg = detect_region(ticker, info)

    currency = info.get('financialCurrency', 'USD')

    # === 分析财年 ===
    fy_analysis = analyze_fiscal_year(info, income_stmt.columns)

    # === 汇率 ===
    fx_rate = 1.0
    if company_cfg['reporting_currency'] != company_cfg['currency']:
        fx_rate = get_fx_rate(
            company_cfg['reporting_currency'],
            company_cfg['currency'],
            proxy
        )

    result = {
        'ticker': ticker,
        'company_name': company_cfg.get('company_name', ''),
        'currency': currency,
        'target_currency': company_cfg['currency'],
        'currency_symbol': company_cfg['currency_symbol'],
        'fx_rate': fx_rate,
        'language': company_cfg['language'],
        'is_china_ops': company_cfg['is_china_ops'],
        'region': company_cfg['region'],
        'fiscal_year_end': info.get('fiscalYearEnd', ''),
        'fiscal_year_analysis': fy_analysis,
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
        # 使用精确映射（基于财年分析），而不是简单的月份判断
        date_str = str(col.date())[:10]
        nat_year = fy_analysis['date_mappings'].get(date_str)
        if nat_year is None:
            # 备选：用旧逻辑
            dt = datetime.strptime(date_str, '%Y-%m-%d')
            nat_year = _map_single_date(dt, fy_analysis['fy_end_month'])

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
                if isinstance(val, float) and math.isnan(val):
                    val = None

                # 货币转换
                if val is not None and fx_rate != 1.0:
                    val = val * fx_rate

                row_data[our_name] = val

        # 如果同年有两条记录（不应该发生，但防一下），保留更近的
        if year_label in result['data']:
            existing = result['data'][year_label]
            # 保留字段更多的那个
            if sum(1 for v in row_data.values() if v is not None) > \
               sum(1 for v in existing.values() if v is not None):
                result['data'][year_label] = row_data
        else:
            result['data'][year_label] = row_data

    # 按年份排序
    result['data'] = dict(sorted(result['data'].items()))

    # 清除完全空的年份
    result['data'] = {
        k: v for k, v in result['data'].items()
        if any(val is not None for val in v.values())
    }

    # 拉取当前市值和PE（也需要货币转换）
    market_cap_raw = info.get('marketCap', None)
    result['current_market_cap'] = market_cap_raw * fx_rate if market_cap_raw else None
    result['current_pe'] = info.get('trailingPE', None)
    result['forward_pe'] = info.get('forwardPE', None)

    current_price_raw = info.get('currentPrice', None)
    result['current_price'] = current_price_raw * fx_rate if current_price_raw else None
    result['current_price_original'] = current_price_raw

    result['shares_outstanding'] = info.get('sharesOutstanding', None)
    result['fifty_two_week_high'] = info.get('fiftyTwoWeekHigh', None)
    result['fifty_two_week_low'] = info.get('fiftyTwoWeekLow', None)

    # 自动选择单位
    if result['data']:
        latest_year = sorted(result['data'].keys())[-1]
        latest_rev = result['data'][latest_year].get('revenue', 0)
        if latest_rev:
            # yfinance 返回的单位已经是原始币种的"百万"
            # 转换后选择合适的显示单位
            unit_info = choose_unit(latest_rev, company_cfg['language'])
            result['unit'] = unit_info['unit']
            result['unit_suffix'] = unit_info['suffix']
            result['unit_label'] = unit_info['label']

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

        facts_url = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"
        resp = requests.get(facts_url, headers=headers, proxies=proxies, timeout=30)
        resp.raise_for_status()
        facts = resp.json()

        revenue_data = {}
        us_gaap = facts.get('facts', {}).get('us-gaap', {})

        revenue_fields = [
            'Revenues',
            'RevenueFromContractWithCustomerExcludingAssessedTax',
            'SalesRevenueNet'
        ]
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

        if i > 0:
            prev_rev = data[years[i-1]].get('revenue')
            if prev_rev and prev_rev != 0 and rev:
                d['revenue_growth'] = round((rev - prev_rev) / prev_rev, 4)

    return data


def main():
    parser = argparse.ArgumentParser(description='拉取个股近5年财务数据')
    parser.add_argument('ticker', help='股票代码 (例: NVDA, 0700.HK)')
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

    # 4. 补充缺失年份
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
        print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
    else:
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
    ccy_sym = result.get('currency_symbol', '$')
    target_ccy = result.get('target_currency', 'USD')
    lang = result.get('language', 'en')

    if lang == 'zh-CN':
        print(f"\n=== 摘要 ===", file=sys.stderr)
        print(f"股票: {args.ticker}", file=sys.stderr)
        print(f"公司: {result.get('company_name', '')}", file=sys.stderr)
        print(f"运营主体: {'中国' if result.get('is_china_ops') else '海外'}", file=sys.stderr)
        print(f"目标货币: {target_ccy}", file=sys.stderr)
        print(f"数据年限: {', '.join(sorted(result['data'].keys()))}", file=sys.stderr)

        fy_analysis = result.get('fiscal_year_analysis', {})
        print(f"\n--- 财年信息 ---", file=sys.stderr)
        print(f"财年结束: {result.get('fiscal_year_end', '未知')}", file=sys.stderr)
        print(f"映射规则: {fy_analysis.get('mapping_rule', '')}", file=sys.stderr)
    else:
        print(f"\n=== Summary ===", file=sys.stderr)
        print(f"Ticker: {args.ticker}", file=sys.stderr)
        print(f"Company: {result.get('company_name', '')}", file=sys.stderr)
        print(f"Target Currency: {target_ccy}", file=sys.stderr)
        print(f"Years: {', '.join(sorted(result['data'].keys()))}", file=sys.stderr)

        fy_analysis = result.get('fiscal_year_analysis', {})
        print(f"\n--- Fiscal Year ---", file=sys.stderr)
        print(f"FY End: {result.get('fiscal_year_end', 'Unknown')}", file=sys.stderr)
        print(f"Mapping: {fy_analysis.get('mapping_rule', '')}", file=sys.stderr)

    if result.get('current_market_cap'):
        print(f"当前市值: {ccy_sym}{result['current_market_cap']/1e9:.1f}B", file=sys.stderr)
    if result.get('current_pe'):
        print(f"当前PE(TTM): {result['current_pe']:.1f}x", file=sys.stderr)
    if result.get('current_price'):
        print(f"当前股价: {ccy_sym}{result['current_price']:.2f}", file=sys.stderr)


if __name__ == '__main__':
    main()
