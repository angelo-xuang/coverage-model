#!/usr/bin/env python3
"""
fetch_segments.py — 拉取个股业务分部数据
用法: python fetch_segments.py <ticker> [--proxy 127.0.0.1:7890]

数据源: SEC EDGAR 10-K (优先) > stockanalysis.com (备选)
"""

import sys
import json
import argparse
import os
import re

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


def fetch_edgar_segments(ticker: str, proxy: str = None) -> dict:
    """
    从 SEC EDGAR 拉取分部收入数据
    返回 {自然年: {分部名: 收入}}
    """
    import requests

    proxies = {}
    if proxy:
        proxies = {'https': f'http://{proxy}', 'http': f'http://{proxy}'}

    headers = {'User-Agent': 'coverage-model@example.com'}

    try:
        # 获取 CIK
        resp = requests.get('https://www.sec.gov/files/company_tickers.json',
                            headers=headers, proxies=proxies, timeout=30)
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

        us_gaap = facts.get('facts', {}).get('us-gaap', {})

        # 搜索分部相关字段
        segment_fields = []
        for key in us_gaap:
            lower = key.lower()
            if ('segment' in lower or 'revenuefromcontract' in lower) and 'segment' in lower:
                segment_fields.append(key)

        # 也搜索具体的分部收入字段
        specific_fields = [
            'RevenueFromContractWithCustomerExcludingAssessedTax',
            'Revenues',
        ]
        for f in specific_fields:
            if f in us_gaap and f not in segment_fields:
                segment_fields.append(f)

        if not segment_fields:
            return {}

        # 尝试从分部信息中提取
        # 查找带 Segment 描述的 revenue 字段
        result = {}
        for field in us_gaap:
            if 'segment' in field.lower() and 'revenue' in field.lower():
                units = us_gaap[field].get('units', {}).get('USD', [])
                for entry in units:
                    if entry.get('form') == '10-K':
                        frame = entry.get('frame', '')
                        if frame.startswith('CY') and 'Q' not in frame:
                            year = frame[2:6] + 'A'
                            if year not in result:
                                result[year] = {}
                            # 用字段名作为分部标识
                            result[year][field] = entry['val']

        return result

    except Exception as e:
        print(f"SEC EDGAR 分部数据获取失败: {e}", file=sys.stderr)
        return {}


def fetch_stockanalysis_segments(ticker: str, proxy: str = None) -> dict:
    """
    从 stockanalysis.com 抓取分部数据（备选方案）
    需要 web_reader MCP 工具，此函数提供数据格式供 AI 调用时使用
    """
    url = f"https://stockanalysis.com/stocks/{ticker.lower()}/financials/?p=quarterly"
    return {
        'url': url,
        'note': '请使用 web_reader 工具抓取此URL，然后手动提取分部数据',
        'expected_format': {
            '2021A': {'分部A': 1000, '分部B': 2000},
            '2022A': {'分部A': 1200, '分部B': 2200},
        }
    }


def main():
    parser = argparse.ArgumentParser(description='拉取个股业务分部数据')
    parser.add_argument('ticker', help='股票代码 (例: NVDA)')
    parser.add_argument('--proxy', default='127.0.0.1:7890', help='代理地址')
    parser.add_argument('--output', choices=['json', 'csv'], default='json')
    parser.add_argument('--source', choices=['edgar', 'web'], default='edgar')
    args = parser.parse_args()

    print(f"正在拉取 {args.ticker} 的业务分部数据...", file=sys.stderr)

    if args.source == 'edgar':
        result = fetch_edgar_segments(args.ticker, args.proxy)
    else:
        result = fetch_stockanalysis_segments(args.ticker, args.proxy)

    if args.output == 'json':
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        if not result:
            print("无数据")
            return
        # CSV
        all_segments = set()
        for year_data in result.values():
            if isinstance(year_data, dict):
                all_segments.update(year_data.keys())

        years = sorted(result.keys())
        print(','.join(['分部'] + years))
        for seg in sorted(all_segments):
            row = [seg]
            for y in years:
                val = result[y].get(seg, '') if isinstance(result[y], dict) else ''
                row.append(str(val) if val else '')
            print(','.join(row))


if __name__ == '__main__':
    main()
