#!/usr/bin/env python3
"""
company_config.py — 公司地区与货币配置

根据 ticker 和公司信息自动判断：
1. 运营主体是否在中国（含港股上市）
2. 使用哪种货币和单位
3. 输出语言（中文 / 英文）
4. 财年定义信息
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

# === 港股 ticker 前缀 → 地区映射 ===
# 格式: ticker前缀 → (交易所, 地区)
HK_PREFIX_MAP = {
    '0': ('HKEX', 'HK'),      # 主板
    '1': ('HKEX', 'HK'),      # 主板
    '2': ('HKEX', 'HK'),      # 主板
    '3': ('HKEX', 'HK'),      # 创业板
    '4': ('HKEX', 'HK'),      # 创业板
    '5': ('HKEX', 'HK'),      # 主板
    '6': ('HKEX', 'HK'),      # 主板
    '7': ('HKEX', 'HK'),      # 创业板
    '8': ('HKEX', 'HK'),      # 创业板
    '9': ('HKEX', 'HK'),      # 创业板
}

# A股前缀
A_SHARE_PREFIXES = ['600', '601', '603', '605', '000', '001', '002', '003', '300', '301']

# 已知的中国运营主体公司（即使在美国上市，运营主体也在中国）
# ticker → 公司名
CHINA_OPS_US_LISTED = {
    'BABA': '阿里巴巴',
    'JD': '京东',
    'PDD': '拼多多',
    'BIDU': '百度',
    'NIO': '蔚来',
    'XPEV': '小鹏汽车',
    'LI': '理想汽车',
    'BILI': '哔哩哔哩',
    'TME': '腾讯音乐',
    'IQ': '爱奇艺',
    'FUTU': '富途',
    'TCEHY': '腾讯控股(ADR)',
    'NTES': '网易',
    'VIPS': '唯品会',
    'ZTO': '中通快递',
    'YMM': '满帮',
    'DIDI': '滴滴',
    'HKBN': '香港宽频',
}

# 已知的国际市场公司（即使总部在亚洲，面向全球）
INTERNATIONAL_OPS = {
    'TSM': ('台积电', '面向全球代工市场'),
    'TSMC': ('台积电', '面向全球代工市场'),
    '2330': ('台积电', '面向全球代工市场'),  # 台股代码
    '005930': ('三星电子', '面向全球市场'),    # 韩股代码
    'SSNLF': ('三星电子', '面向全球市场'),     # 美股OTC
    '006400': ('三星电子', '面向全球市场'),
    'ASML': ('ASML', '全球半导体设备'),
    'ARM': ('ARM', '全球IP授权'),
}


def detect_region(ticker: str, yfinance_info: dict = None) -> dict:
    """
    检测公司的运营地区和货币配置。

    返回:
    {
        'ticker': str,
        'company_name': str,        # 公司名
        'exchange': str,            # 交易所
        'region': str,              # 'CN', 'HK', 'TW', 'KR', 'US', 'EU', 'JP', etc.
        'is_china_ops': bool,       # 运营主体是否在中国
        'currency': str,            # 'CNY', 'USD', 'EUR', etc.
        'currency_symbol': str,     # '¥', '$', '€', etc.
        'reporting_currency': str,  # 财报原始币种
        'unit': str,                # 'Million' / 'Billion'
        'unit_label': str,          # '百万' / '十亿' / 'Million' / 'Billion'
        'language': str,            # 'zh-CN' / 'en'
        'unit_suffix': str,         # 'M' / 'B'
        'fx_rate_to_target': float, # 原始币种 → 目标币种的汇率
        'fiscal_year_end_month': int, # 财年结束月份
        'fiscal_year_note': str,    # 财年说明
    }
    """
    ticker_upper = ticker.upper().strip()

    # 默认值
    config = {
        'ticker': ticker_upper,
        'company_name': '',
        'exchange': '',
        'region': 'US',
        'is_china_ops': False,
        'currency': 'USD',
        'currency_symbol': '$',
        'reporting_currency': 'USD',
        'unit': 'Billion',
        'unit_suffix': 'B',
        'language': 'en',
        'fx_rate_to_target': 1.0,
        'fiscal_year_end_month': 12,
        'fiscal_year_note': '',
    }

    # --- Step 1: 根据 ticker 格式初步判断 ---

    # 港股: 4-5位纯数字，带 .HK 后缀
    if ticker_upper.endswith('.HK') or (ticker_upper.isdigit() and 4 <= len(ticker_upper) <= 5):
        config['exchange'] = 'HKEX'
        config['region'] = 'HK'
        config['is_china_ops'] = True
        config['reporting_currency'] = 'HKD'
        config['currency'] = 'CNY'
        config['currency_symbol'] = '¥'
        config['language'] = 'zh-CN'

    # A股: 6位纯数字
    elif ticker_upper.isdigit() and len(ticker_upper) == 6:
        prefix = ticker_upper[:3]
        if prefix in A_SHARE_PREFIXES:
            config['exchange'] = 'SSE' if prefix in ['600', '601', '603', '605'] else 'SZSE'
            if prefix in ['300', '301']:
                config['exchange'] = 'SZSE-STAR'
            config['region'] = 'CN'
            config['is_china_ops'] = True
            config['reporting_currency'] = 'CNY'
            config['currency'] = 'CNY'
            config['currency_symbol'] = '¥'
            config['language'] = 'zh-CN'

    # 台股: 4位纯数字
    elif ticker_upper.isdigit() and len(ticker_upper) == 4:
        config['exchange'] = 'TWSE'
        config['region'] = 'TW'
        config['reporting_currency'] = 'TWD'
        config['currency'] = 'USD'
        config['currency_symbol'] = '$'
        config['language'] = 'en'

    # 韩股: 6位纯数字
    elif ticker_upper.isdigit() and len(ticker_upper) == 6:
        config['exchange'] = 'KRX'
        config['region'] = 'KR'
        config['reporting_currency'] = 'KRW'
        config['currency'] = 'USD'
        config['currency_symbol'] = '$'
        config['language'] = 'en'

    # --- Step 2: 检查特殊列表 ---

    if ticker_upper in CHINA_OPS_US_LISTED:
        config['is_china_ops'] = True
        config['company_name'] = CHINA_OPS_US_LISTED[ticker_upper]
        config['region'] = 'CN'
        config['reporting_currency'] = 'USD'  # 美股上市通常用USD报告
        config['currency'] = 'CNY'
        config['currency_symbol'] = '¥'
        config['language'] = 'zh-CN'

    if ticker_upper in INTERNATIONAL_OPS:
        name, note = INTERNATIONAL_OPS[ticker_upper]
        config['company_name'] = name
        config['is_china_ops'] = False
        config['currency'] = 'USD'
        config['currency_symbol'] = '$'
        config['language'] = 'en'

    # --- Step 3: 用 yfinance info 补充和修正 ---

    if yfinance_info:
        info = yfinance_info

        # 公司名
        if not config['company_name']:
            config['company_name'] = info.get('longName', info.get('shortName', ''))

        # 财报原始币种
        reporting_ccy = info.get('financialCurrency', '')
        if reporting_ccy:
            config['reporting_currency'] = reporting_ccy

        # 财年结束日期
        fy_end = info.get('fiscalYearEnd', '')
        if fy_end:
            try:
                # 格式通常是 "MM-DD"
                month = int(fy_end.split('-')[0])
                config['fiscal_year_end_month'] = month
                config['fiscal_year_note'] = _fiscal_year_note(month, ticker_upper)
            except (ValueError, IndexError):
                pass

        # 通过country/地址判断
        country = info.get('country', '')
        state = info.get('state', '')

        # 如果 yfinance 返回了中国地址，且之前没标记为中国
        if country in ('China', 'Hong Kong', 'Taiwan') and not config['is_china_ops']:
            if country == 'China':
                config['is_china_ops'] = True
                config['region'] = 'CN'
                config['language'] = 'zh-CN'
            elif country == 'Hong Kong':
                config['is_china_ops'] = True
                config['region'] = 'HK'
                config['language'] = 'zh-CN'

        # 交易所
        exchange_val = info.get('exchange', '')
        if exchange_val and not config['exchange']:
            config['exchange'] = exchange_val

    # --- Step 4: 确定单位和语言 ---
    if config['language'] == 'zh-CN':
        config['unit_label'] = '十亿' if config['unit'] == 'Billion' else '百万'
    else:
        config['unit_label'] = config['unit']

    return config


def _fiscal_year_note(fy_end_month: int, ticker: str) -> str:
    """生成财年说明"""
    month_names_zh = {
        1: '1月', 2: '2月', 3: '3月', 4: '4月', 5: '5月', 6: '6月',
        7: '7月', 8: '8月', 9: '9月', 10: '10月', 11: '11月', 12: '12月',
    }

    known_fy = {
        'NVDA': 'NVIDIA财年结束于1月底（通常为1月最后一个周日），FY2026对应自然年2025年。'
                 '因为FY结束日在1月，该FY的收入大部分产生在前一个自然年。',
        'AAPL': 'Apple财年结束于9月底，FY2024对应自然年2024年（10月-9月）。',
        'MSFT': 'Microsoft财年结束于6月底，FY2025对应自然年2024年（7月-6月）。',
        'GOOGL': 'Google财年结束于12月底，与自然年一致。',
        'TSM': '台积电财年结束于12月底，与自然年一致。',
    }

    if ticker.upper() in known_fy:
        return known_fy[ticker.upper()]

    if fy_end_month == 12:
        return f'财年结束于{month_names_zh[12]}，与自然年一致。'
    elif fy_end_month <= 6:
        return (f'财年结束于{month_names_zh[fy_end_month]}，FY标签对应上一个自然年。'
                f'例如FY2026 ≈ 自然年2025。'
                f'需要根据实际财报日期精确映射。')
    else:
        return (f'财年结束于{month_names_zh[fy_end_month]}，FY标签通常与当年自然年对应。'
                f'但需核实具体映射关系。')


def get_fx_rate(from_ccy: str, to_ccy: str, proxy: str = None) -> float:
    """
    获取汇率（从 from_ccy 到 to_ccy）
    优先用 yfinance，备选硬编码常见汇率
    """
    import os

    # 同币种
    if from_ccy == to_ccy:
        return 1.0

    # 常见汇率硬编码（会在运行时尝试更新）
    HARDCODED_RATES = {
        ('HKD', 'CNY'): 0.92,
        ('HKD', 'USD'): 0.128,
        ('TWD', 'USD'): 0.033,
        ('KRW', 'USD'): 0.00073,
        ('EUR', 'USD'): 1.10,
        ('JPY', 'USD'): 0.0064,
        ('GBP', 'USD'): 1.27,
        ('USD', 'CNY'): 7.25,
        ('CNY', 'USD'): 0.138,
    }

    key = (from_ccy, to_ccy)
    if key in HARDCODED_RATES:
        rate = HARDCODED_RATES[key]
    else:
        # 尝试反向
        reverse = (to_ccy, from_ccy)
        if reverse in HARDCODED_RATES:
            rate = 1.0 / HARDCODED_RATES[reverse]
        else:
            rate = 1.0

    # 尝试用 yfinance 获取实时汇率
    try:
        if proxy:
            os.environ['HTTPS_PROXY'] = f'http://{proxy}'
        import yfinance as yf
        fx_ticker = f"{from_ccy}{to_ccy}=X"
        fx_data = yf.Ticker(fx_ticker)
        hist = fx_data.history(period="5d")
        if not hist.empty:
            live_rate = hist['Close'].iloc[-1]
            if live_rate and live_rate > 0:
                rate = float(live_rate)
    except Exception:
        pass

    return rate


def choose_unit(total_value_in_target_ccy: float, language: str = 'zh-CN') -> dict:
    """
    根据数值大小选择合适的单位

    Args:
        total_value: 以目标币种计的数值（如收入、市值等）
        language: 语言

    Returns:
        {'unit': 'Billion'/'Million', 'suffix': 'B'/'M', 'label': '十亿'/'百万'/'Billion'/'Million'}
    """
    # 如果数值 > 1000（以百万计），用 Billion
    if abs(total_value_in_target_ccy) >= 1000:
        unit = 'Billion'
        suffix = 'B'
    else:
        unit = 'Million'
        suffix = 'M'

    if language == 'zh-CN':
        label = '十亿' if unit == 'Billion' else '百万'
    else:
        label = unit

    return {'unit': unit, 'suffix': suffix, 'label': label}


def format_currency_label(config: dict) -> str:
    """生成货币单位标签，如 '¥B' 或 '$M' 或 '百万人民币' 等"""
    if config['language'] == 'zh-CN':
        return f"{config['currency_symbol']}{config.get('unit_suffix', 'B')}"
    return f"{config['currency_symbol']}{config.get('unit_suffix', 'B')}"


if __name__ == '__main__':
    import json
    import argparse

    parser = argparse.ArgumentParser(description='公司地区与货币配置检测')
    parser.add_argument('ticker', help='股票代码')
    parser.add_argument('--proxy', default='127.0.0.1:7890')
    args = parser.parse_args()

    # 尝试获取 yfinance info
    yf_info = None
    try:
        import os
        if args.proxy:
            os.environ['HTTPS_PROXY'] = f'http://{args.proxy}'
        import yfinance as yf
        stock = yf.Ticker(args.ticker)
        yf_info = stock.info
    except Exception as e:
        print(f"yfinance info 获取失败: {e}", file=sys.stderr)

    config = detect_region(args.ticker, yf_info)

    # 获取汇率
    if config['reporting_currency'] != config['currency']:
        fx = get_fx_rate(config['reporting_currency'], config['currency'], args.proxy)
        config['fx_rate_to_target'] = fx

    print(json.dumps(config, ensure_ascii=False, indent=2))
