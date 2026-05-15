#!/usr/bin/env python3
"""
googl_model_data.py — Google (GOOGL) 估值模型数据准备
分部预测版：每个业务线独立的 TAM × 市占率 = 收入
"""

import json
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 分部假设：每个分部独立的 TAM 和市占率
# ============================================================

SEGMENT_ASSUMPTIONS = {
    'bear': {
        'search': {
            'tam_2025': 247.0,           # $B - 搜索广告市场 (Search $224.5B / 91%市占率)
            'tam_cagr_1_3': 0.08,        # eMarketer: 数字广告增速首次个位数
            'tam_cagr_4_6': 0.06,
            'tam_cagr_7_10': 0.04,
            'share_2025': 0.91,          # Statista 2025
            'share_2035': 0.85,          # AI搜索蚕食
        },
        'youtube': {
            'tam_2025': 135.0,           # $B - 视频广告市场 (YouTube $40.4B / 30%市占率估算)
            'tam_cagr_1_3': 0.10,
            'tam_cagr_4_6': 0.08,
            'tam_cagr_7_10': 0.05,
            'share_2025': 0.30,          # 估算
            'share_2035': 0.25,          # 短视频竞争(TikTok)
        },
        'cloud': {
            'tam_2025': 451.0,           # $B - 云基础设施 (Synergy Research Q2'25)
            'tam_cagr_1_3': 0.16,        # Gartner: 2026公有云增速21.3%
            'tam_cagr_4_6': 0.12,
            'tam_cagr_7_10': 0.08,
            'share_2025': 0.13,          # Synergy Q2'25
            'share_2035': 0.12,          # 竞争加剧
        },
        'network': {
            # Network是萎缩业务，直接预测收入而非TAM×市占率
            'revenue_2025': 29.8,        # $B - 2025A实际
            'annual_decline': 0.05,      # 每年萎缩5%（隐私政策压制）
        },
        'subscriptions': {
            'tam_2025': 320.0,           # $B - 订阅+设备市场估算
            'tam_cagr_1_3': 0.08,        # Grand View: streaming CAGR 13.2%
            'tam_cagr_4_6': 0.06,
            'tam_cagr_7_10': 0.04,
            'share_2025': 0.15,          # 估算
            'share_2035': 0.12,
        },
        # 通用参数
        'gm_start': 0.597,
        'gm_end': 0.580,
        'expense_start': 0.277,
        'expense_end': 0.290,
        'tax_rate': 0.17,
        'pe_start': 22,
        'pe_end': 15,
        'shares_m': 5824,
    },
    'base': {
        'search': {
            'tam_2025': 247.0,
            'tam_cagr_1_3': 0.10,        # eMarketer: 数字广告增速10%
            'tam_cagr_4_6': 0.08,
            'tam_cagr_7_10': 0.05,
            'share_2025': 0.91,
            'share_2035': 0.88,          # AI搜索蚕食但Google应对得当
        },
        'youtube': {
            'tam_2025': 135.0,
            'tam_cagr_1_3': 0.12,        # 视频广告增速>整体广告
            'tam_cagr_4_6': 0.10,
            'tam_cagr_7_10': 0.06,
            'share_2025': 0.30,
            'share_2035': 0.28,
        },
        'cloud': {
            'tam_2025': 451.0,
            'tam_cagr_1_3': 0.21,        # Gartner: 2026公有云增速21.3%
            'tam_cagr_4_6': 0.15,
            'tam_cagr_7_10': 0.10,
            'share_2025': 0.13,
            'share_2035': 0.17,          # GCP从10%(2022)→13%(2025)→17%(2035)
        },
        'network': {
            'revenue_2025': 29.8,
            'annual_decline': 0.03,      # 每年萎缩3%
        },
        'subscriptions': {
            'tam_2025': 320.0,
            'tam_cagr_1_3': 0.13,        # Grand View: streaming CAGR 13.2%
            'tam_cagr_4_6': 0.10,
            'tam_cagr_7_10': 0.06,
            'share_2025': 0.15,
            'share_2035': 0.18,          # YouTube Premium增长
        },
        'gm_start': 0.597,
        'gm_end': 0.615,                 # +1.8pp/10年，Cloud扭亏+AI提升搜索变现
        'expense_start': 0.277,
        'expense_end': 0.260,            # -1.7pp/10年
        'tax_rate': 0.15,
        'pe_start': 27,
        'pe_end': 20,
        'shares_m': 5824,
    },
    'bull': {
        'search': {
            'tam_2025': 247.0,
            'tam_cagr_1_3': 0.12,
            'tam_cagr_4_6': 0.10,
            'tam_cagr_7_10': 0.06,
            'share_2025': 0.91,
            'share_2035': 0.90,          # AI搜索但Google维持优势
        },
        'youtube': {
            'tam_2025': 135.0,
            'tam_cagr_1_3': 0.15,
            'tam_cagr_4_6': 0.12,
            'tam_cagr_7_10': 0.08,
            'share_2025': 0.30,
            'share_2035': 0.32,          # YouTube Shorts成功
        },
        'cloud': {
            'tam_2025': 451.0,
            'tam_cagr_1_3': 0.25,        # AI爆发，云增速超预期
            'tam_cagr_4_6': 0.18,
            'tam_cagr_7_10': 0.12,
            'share_2025': 0.13,
            'share_2035': 0.20,          # GCP突破到20%
        },
        'network': {
            'revenue_2025': 29.8,
            'annual_decline': 0.01,      # 每年仅萎缩1%
        },
        'subscriptions': {
            'tam_2025': 320.0,
            'tam_cagr_1_3': 0.15,
            'tam_cagr_4_6': 0.12,
            'tam_cagr_7_10': 0.08,
            'share_2025': 0.15,
            'share_2035': 0.20,
        },
        'gm_start': 0.597,
        'gm_end': 0.640,                 # +4.3pp/10年
        'expense_start': 0.277,
        'expense_end': 0.245,            # -3.2pp/10年
        'tax_rate': 0.14,
        'pe_start': 33,
        'pe_end': 25,
        'shares_m': 5824,
    },
}

# 历史财务数据 (单位: $M)
HIST_DATA = {
    '2021A': {'revenue': 257637},
    '2022A': {'revenue': 282836, 'cogs': 126203, 'gross_profit': 156633,
              'rd': 39500, 'sga': 26598, 'ebit': 74988, 'net_income': 59972,
              'gross_margin': 0.5538, 'ebit_margin': 0.2651, 'net_margin': 0.2120,
              'rd_ratio': 0.1397, 'sga_ratio': 0.0940},
    '2023A': {'revenue': 307394, 'cogs': 133452, 'gross_profit': 173942,
              'rd': 45364, 'sga': 24870, 'ebit': 84233, 'net_income': 73795,
              'gross_margin': 0.5658, 'ebit_margin': 0.2740, 'net_margin': 0.2401,
              'rd_ratio': 0.1476, 'sga_ratio': 0.0809, 'revenue_growth': 0.0868},
    '2024A': {'revenue': 350018, 'cogs': 146318, 'gross_profit': 203700,
              'rd': 45427, 'sga': 28400, 'ebit': 112230, 'net_income': 100118,
              'gross_margin': 0.5819, 'ebit_margin': 0.3206, 'net_margin': 0.2860,
              'rd_ratio': 0.1298, 'sga_ratio': 0.0811, 'revenue_growth': 0.1387},
    '2025A': {'revenue': 402779, 'cogs': 162322, 'gross_profit': 240457,
              'rd': 49600, 'sga': 29500, 'ebit': 128940, 'net_income': 132190,
              'gross_margin': 0.5969, 'ebit_margin': 0.3200, 'net_margin': 0.3282,
              'rd_ratio': 0.1231, 'sga_ratio': 0.0732, 'revenue_growth': 0.1507,
              # 分部收入 (2025A估算)
              'search_revenue': 224500,      # Search & Other
              'youtube_revenue': 40400,      # YouTube Ads
              'cloud_revenue': 58700,        # Google Cloud
              'network_revenue': 29800,      # Network (AdSense)
              'subscriptions_revenue': 48000, # Subscriptions & Devices
              },
}

# 估值数据
VAL_DATA = {
    'current_market_cap': 4692900000000,
    'current_price': 387.35,
    'current_pe': 29.6,
    'forward_pe': 26.8,
    'shares_outstanding': 5824000000,
    'fifty_two_week_high': 402.00,
    'fifty_two_week_low': 159.61,
}


def get_cagr_for_year(t, cagr_1_3, cagr_4_6, cagr_7_10):
    """获取第t年适用的CAGR"""
    if t <= 3:
        return cagr_1_3
    elif t <= 6:
        return cagr_4_6
    else:
        return cagr_7_10


def project_segment(segment_name, seg_params, t, prev_tam=None, prev_rev=None):
    """
    预测单个分部第t年的收入

    对于TAM×市占率的分部(search/youtube/cloud/subscriptions):
        tam_t = tam_prev × (1 + cagr)
        share_t = share_start + (share_end - share_start) × (t-1)/(n_years-1)
        revenue = tam_t × share_t

    对于萎缩业务(network):
        revenue_t = revenue_prev × (1 - annual_decline)
    """
    if segment_name == 'network':
        # Network: 直接预测收入萎缩
        if t == 1:
            rev = seg_params['revenue_2025'] * (1 - seg_params['annual_decline'])
        else:
            rev = prev_rev * (1 - seg_params['annual_decline'])
        return {'revenue': rev, 'tam': None, 'share': None}
    else:
        # 其他分部: TAM × 市占率
        cagr = get_cagr_for_year(t,
                                 seg_params['tam_cagr_1_3'],
                                 seg_params['tam_cagr_4_6'],
                                 seg_params['tam_cagr_7_10'])
        if t == 1:
            tam = seg_params['tam_2025'] * (1 + cagr)
        else:
            tam = prev_tam * (1 + cagr)

        share_start = seg_params['share_2025']
        share_end = seg_params['share_2035']
        # 市占率在第1年开始变化（t=1时已经是预测年）
        share = share_start + (share_end - share_start) * t / 10

        revenue = tam * share
        return {'revenue': revenue, 'tam': tam, 'share': share}


def project_scenario(name: str, params: dict, hist_data: dict, n_years: int = 10):
    """对单个情景做完整分部预测"""
    last_year = max(int(k.replace('A', '')) for k in hist_data)
    last_rev_M = hist_data[f'{last_year}A']['revenue']  # $M
    last_rev_B = last_rev_M / 1000  # $B

    # 分部参数
    search_params = params['search']
    youtube_params = params['youtube']
    cloud_params = params['cloud']
    network_params = params['network']
    subscriptions_params = params['subscriptions']

    # 通用参数
    gm_start = params['gm_start']
    gm_end = params['gm_end']
    exp_start = params['expense_start']
    exp_end = params['expense_end']
    tax_rate = params['tax_rate']
    pe_start = params['pe_start']
    pe_end = params['pe_end']
    shares_m = params['shares_m']

    rows = {y: {} for y in range(1, n_years + 1)}

    # 分部历史状态（用于递推）
    prev_search_tam = None
    prev_youtube_tam = None
    prev_cloud_tam = None
    prev_subscriptions_tam = None
    prev_network_rev = None

    for t in range(1, n_years + 1):
        # 预测各分部
        search_result = project_segment('search', search_params, t, prev_search_tam)
        youtube_result = project_segment('youtube', youtube_params, t, prev_youtube_tam)
        cloud_result = project_segment('cloud', cloud_params, t, prev_cloud_tam)
        network_result = project_segment('network', network_params, t, prev_rev=prev_network_rev)
        subscriptions_result = project_segment('subscriptions', subscriptions_params, t, prev_subscriptions_tam)

        # 汇总收入
        total_rev = (search_result['revenue'] +
                     youtube_result['revenue'] +
                     cloud_result['revenue'] +
                     network_result['revenue'] +
                     subscriptions_result['revenue'])

        # 更新历史状态
        prev_search_tam = search_result['tam']
        prev_youtube_tam = youtube_result['tam']
        prev_cloud_tam = cloud_result['tam']
        prev_subscriptions_tam = subscriptions_result['tam']
        prev_network_rev = network_result['revenue']

        # 毛利率、费用率线性插值
        gm = gm_start + (gm_end - gm_start) * (t - 1) / (n_years - 1)
        gross_profit = total_rev * gm
        exp = exp_start + (exp_end - exp_start) * (t - 1) / (n_years - 1)
        opex = total_rev * exp
        ebit = gross_profit - opex
        net_income = ebit * (1 - tax_rate)

        # 收入增速
        if t == 1:
            rev_growth = (total_rev - last_rev_B) / last_rev_B
        else:
            prev_rev = rows[t - 1]['total_rev']
            rev_growth = (total_rev - prev_rev) / prev_rev if prev_rev else 0

        # PE线性衰减
        pe = pe_start + (pe_end - pe_start) * (t - 1) / (n_years - 1)

        rows[t] = {
            # 分部明细
            'search_tam': search_result['tam'],
            'search_share': search_result['share'],
            'search_rev': search_result['revenue'],
            'youtube_tam': youtube_result['tam'],
            'youtube_share': youtube_result['share'],
            'youtube_rev': youtube_result['revenue'],
            'cloud_tam': cloud_result['tam'],
            'cloud_share': cloud_result['share'],
            'cloud_rev': cloud_result['revenue'],
            'network_rev': network_result['revenue'],
            'subscriptions_tam': subscriptions_result['tam'],
            'subscriptions_share': subscriptions_result['share'],
            'subscriptions_rev': subscriptions_result['revenue'],
            # 汇总
            'total_rev': total_rev,
            'rev_growth': rev_growth,
            'gm': gm,
            'gross_profit': gross_profit,
            'exp': exp,
            'opex': opex,
            'ebit': ebit,
            'tax_rate': tax_rate,
            'net_income': net_income,
            'ebit_margin': ebit / total_rev if total_rev else 0,
            'net_margin': net_income / total_rev if total_rev else 0,
            'pe': pe,
        }

    # 终值计算
    terminal_ni = rows[n_years]['net_income']
    terminal_pe = rows[n_years]['pe']
    mkt_cap = terminal_ni * terminal_pe
    price = mkt_cap * 1000 / shares_m
    annual_return = (price / VAL_DATA['current_price']) ** (1 / n_years) - 1

    # 收入CAGR
    cagr = (total_rev / last_rev_B) ** (1 / n_years) - 1

    return {
        'name': name,
        'rows': rows,
        'terminal': {
            'year': last_year + n_years,
            'rev_B': total_rev,
            'gm': gm,
            'ebit_margin': ebit / total_rev if total_rev else 0,
            'net_income_B': net_income,
            'mkt_cap_B': mkt_cap,
            'price': price,
            'cagr': cagr,
            'annual_return': annual_return,
            'pe': terminal_pe,
        }
    }


def sanity_checks(scenarios: list, hist_data: dict):
    """三层合理性检验"""
    results = []
    n_years = 10

    # === 第一层：内部逻辑 ===
    for s in scenarios:
        name = s['name'].upper()
        rows = s['rows']

        # 收入增速衰减检查
        growths = [rows[t]['rev_growth'] for t in range(1, n_years + 1)]
        accel_count = 0
        for i in range(1, len(growths)):
            if growths[i] > growths[i-1]:
                accel_count += 1
            else:
                accel_count = 0
        growth_decay_ok = accel_count < 2

        max_growth = max(growths)
        growth_abs_ok = max_growth < 0.60

        gms = [rows[t]['gm'] for t in range(1, n_years + 1)]
        gm_ok = all(0.30 < g < 0.80 for g in gms)

        ebm = [rows[t]['ebit_margin'] for t in range(1, n_years + 1)]
        ebit_ok = all(0 < e < 0.55 for e in ebm)

        results.append({
            'name': f'[{name}] 收入增速衰减',
            'status': '✓ 通过' if growth_decay_ok else '⚠ 注意',
        })
        results.append({
            'name': f'[{name}] 毛利率范围',
            'status': '✓ 通过' if gm_ok else '✗ 不通过',
        })
        results.append({
            'name': f'[{name}] EBIT margin 合理',
            'status': '✓ 通过' if ebit_ok else '✗ 不通过',
        })

    # === 第二层：跨情景排序 ===
    bear_t = scenarios[0]['terminal']
    base_t = scenarios[1]['terminal']
    bull_t = scenarios[2]['terminal']

    order_ok = bull_t['rev_B'] > base_t['rev_B'] > bear_t['rev_B']
    ni_order_ok = bull_t['net_income_B'] > base_t['net_income_B'] > bear_t['net_income_B']

    bull_base_gap = (bull_t['price'] - base_t['price']) / base_t['price']
    base_bear_gap = (base_t['price'] - bear_t['price']) / bear_t['price']
    gap_ok = 0.3 < bull_base_gap < 2.0 and 0.3 < base_bear_gap < 2.0

    results.append({
        'name': '终年收入排序 Bull>Base>Bear',
        'bear': f"${bear_t['rev_B']:.0f}B",
        'base': f"${base_t['rev_B']:.0f}B",
        'bull': f"${bull_t['rev_B']:.0f}B",
        'status': '✓ 通过' if order_ok else '✗ 不通过',
    })
    results.append({
        'name': '终年净利排序 Bull>Base>Bear',
        'bear': f"${bear_t['net_income_B']:.0f}B",
        'base': f"${base_t['net_income_B']:.0f}B",
        'bull': f"${bull_t['net_income_B']:.0f}B",
        'status': '✓ 通过' if ni_order_ok else '✗ 不通过',
    })
    results.append({
        'name': f'情景差距合理性 (B/B={base_bear_gap:.0%}, Bu/B={bull_base_gap:.0%})',
        'status': '✓ 通过' if gap_ok else '⚠ 差距偏大',
    })

    # === 第三层：外部基准 ===
    base_cagr = base_t['cagr']
    results.append({
        'name': f'Base CAGR ({base_cagr:.1%}) vs MSFT历史 (~12%)',
        'status': '✓ 合理' if 0.05 < base_cagr < 0.20 else '⚠ 偏离',
    })

    return results


def main():
    scenarios = []
    for name in ['bear', 'base', 'bull']:
        proj = project_scenario(name, SEGMENT_ASSUMPTIONS[name], HIST_DATA)
        scenarios.append(proj)

        t = proj['terminal']
        print(f"\n{'='*60}")
        print(f"  {name.upper()} CASE — 终年 {t['year']}E")
        print(f"{'='*60}")
        print(f"  收入:       ${t['rev_B']:>10,.1f}B  (CAGR: {t['cagr']:.1%})")
        print(f"  毛利率:     {t['gm']:>10.1%}")
        print(f"  EBIT margin:{t['ebit_margin']:>10.1%}")
        print(f"  净利润:     ${t['net_income_B']:>10,.1f}B")
        print(f"  终年PE:     {t['pe']:>10.1f}x")
        print(f"  隐含市值:   ${t['mkt_cap_B']:>10,.1f}B")
        print(f"  隐含股价:   ${t['price']:>10,.1f}")
        print(f"  vs 当前:    {(t['price']/387.35 - 1):>10.1%}")
        print(f"  年化回报:   {t['annual_return']:>10.1%}")

    # Base case 分部明细
    base_rows = scenarios[1]['rows']
    print(f"\n{'='*100}")
    print(f"  BASE CASE — 分部逐年明细")
    print(f"{'='*100}")
    header = (f"{'Year':>6} "
              f"{'Search':>12} {'Search%':>7} "
              f"{'YouTube':>12} {'YT%':>7} "
              f"{'Cloud':>12} {'Cloud%':>7} "
              f"{'Network':>10} "
              f"{'Subs':>10} "
              f"{'Total':>12} {'Gr':>6} {'GM':>6} {'NI':>10} {'PE':>5}")
    print(f"  {header}")
    print(f"  {'-'*96}")

    for t in range(1, 11):
        r = base_rows[t]
        year = 2025 + t

        # 分部占比
        search_pct = r['search_rev'] / r['total_rev'] * 100 if r['total_rev'] else 0
        yt_pct = r['youtube_rev'] / r['total_rev'] * 100 if r['total_rev'] else 0
        cloud_pct = r['cloud_rev'] / r['total_rev'] * 100 if r['total_rev'] else 0

        print(f"  {year}E "
              f"${r['search_rev']:>10,.1f}B {r['search_share']*100:>5.1f}% "
              f"${r['youtube_rev']:>10,.1f}B {r['youtube_share']*100:>5.1f}% "
              f"${r['cloud_rev']:>10,.1f}B {r['cloud_share']*100:>5.1f}% "
              f"${r['network_rev']:>8,.1f}B "
              f"${r['subscriptions_rev']:>8,.1f}B "
              f"${r['total_rev']:>10,.1f}B {r['rev_growth']*100:>5.1f}% "
              f"{r['gm']*100:>5.1f}% "
              f"${r['net_income']:>8,.1f}B "
              f"{r['pe']:>4.1f}x")

    # 合理性检验
    checks = sanity_checks(scenarios, HIST_DATA)
    print(f"\n{'='*60}")
    print(f"  合理性检验")
    print(f"{'='*60}")
    for c in checks:
        vals = ' | '.join(filter(None, [c.get('bear', ''), c.get('base', ''), c.get('bull', '')]))
        if vals:
            print(f"  {c['name']:<40} {vals} | {c['status']}")
        else:
            print(f"  {c['name']:<40} {c['status']}")

    # 分部假设来源注释
    sources = {
        'search_tam': (
            "搜索广告市场TAM = $247B\n"
            "推导: Search & Other收入 $224.5B / 市占率91% (Statista 2025)\n"
            "= $247B\n"
            "数据来源: Statista Search Engine Market Share 2025"
        ),
        'search_share': (
            "起始市占率91% = Statista 2025实际\n"
            "终年市占率:\n"
            "  Base 88%: AI搜索蚕食，参考Intel CPU市占率从90%→70%花8年\n"
            "  Bull 90%: Google在AI搜索应对得当\n"
            "  Bear 85%: AI颠覆加速"
        ),
        'search_tam_cagr': (
            "Base 10% = eMarketer 2025数字广告增速预测\n"
            "eMarketer: 全球数字广告增速首次降至个位数(2025)\n"
            "Bull 12% / Bear 8%"
        ),
        'youtube_tam': (
            "视频广告市场TAM = $135B\n"
            "推导: YouTube Ads收入 $40.4B / 市占率30% (估算)\n"
            "= $135B\n"
            "市占率估算方法: YouTube在视频广告市场的份额（vs TikTok, Meta等）"
        ),
        'youtube_share': (
            "起始市占率30% = 估算\n"
            "终年市占率:\n"
            "  Base 28%: 短视频竞争(TikTok)持续\n"
            "  Bull 32%: YouTube Shorts成功\n"
            "  Bear 25%: TikTok抢占更多份额"
        ),
        'cloud_tam': (
            "云基础设施市场TAM = $451B\n"
            "数据来源: Synergy Research Group Q2 2025\n"
            "全球云基础设施市场规模（含IaaS+PaaS）"
        ),
        'cloud_share': (
            "起始市占率13% = Synergy Research Q2 2025实际\n"
            "历史轨迹: GCP从10%(2022)→13%(2025)，每年+1pp\n"
            "终年市占率:\n"
            "  Base 17%: 每年+0.4pp，竞争加剧\n"
            "  Bull 20%: GCP突破\n"
            "  Bear 12%: 竞争失利"
        ),
        'cloud_tam_cagr': (
            "Base 21% (1-3年) = Gartner 2026公有云增速预测21.3%\n"
            "Gartner: GenAI驱动云支出增长\n"
            "Bull 25% / Bear 16%"
        ),
        'network': (
            "Network(AdSense)为萎缩业务\n"
            "2025A收入 $29.8B\n"
            "年萎缩率: Base 3% / Bear 5% / Bull 1%\n"
            "原因: 隐私政策(Cookie限制)、第三方广告网络整体萎缩"
        ),
        'subscriptions_tam': (
            "订阅+设备市场TAM = $320B\n"
            "推导: Subscriptions收入 $48B / 市占率15% (估算)\n"
            "= $320B\n"
            "含: YouTube Premium、Google One、Pixel设备等"
        ),
        'subscriptions_share': (
            "起始市占率15% = 估算\n"
            "终年市占率:\n"
            "  Base 18%: YouTube Premium增长\n"
            "  Bull 20%: 订阅业务突破\n"
            "  Bear 12%: 竞争加剧"
        ),
        'gm_start': (
            "59.7% = 2025A实际值\n"
            "历史轨迹: 55.4%(2022)→56.6%(2023)→58.2%(2024)→59.7%(2025)\n"
            "年均+1.1pp，主因Cloud扭亏"
        ),
        'gm_end': (
            "Base 61.5% (+1.8pp/10年):\n"
            "扩张支撑:\n"
            "  (1) Cloud扭亏: GCP 2023首次季度盈利\n"
            "  (2) AI提升搜索变现: targeting精准→CPM↑\n"
            "压制因素:\n"
            "  (1) AI CapEx折旧进COGS\n"
            "  (2) Cloud占比↑但利润率低\n"
            "  (3) 硬件(Pixel)增长，硬件GM 30-40%远低于搜索"
        ),
        'expense_start': (
            "27.7% = 2025A实际\n"
            "= (收入 - COGS - EBIT) / 收入\n"
            "= R&D $49.6B(12.3%) + SGA $29.5B(7.3%) + 其他"
        ),
        'expense_end': (
            "Base 26.0% (-1.7pp/10年):\n"
            "参考: META费用率从45%(2022)降到30%(2024)\n"
            "限制因素: AI人才成本高、Gemini需持续R&D"
        ),
        'tax_rate': (
            "Base 15%:\n"
            "GOOGL历史有效税率 ~16% (2024A 10-K)\n"
            "影响因素: 海外收入税率较低、R&D税收抵免"
        ),
        'pe_start': (
            "PE起始值:\n"
            "Base 27x = GOOGL 5年PE中位27.4 (yfinance)\n"
            "当前TTM 29.6x, Forward 26.8x\n"
            "Bull 33x / Bear 22x"
        ),
        'pe_end': (
            "PE终年值(增速衰减后):\n"
            "Base 20x: 增速降至~6%时合理PE\n"
            "参考: MSFT PE从35x(高增长)→25x→20x(成熟)\n"
            "Bull 25x / Bear 15x"
        ),
        'shares_m': (
            "5,824M = yfinance sharesOutstanding\n"
            "假设10年不变(保守: GOOGL回购力度大，实际可能降)"
        ),
    }

    # 输出 JSON
    output = {
        'hist_data': HIST_DATA,
        'val_data': VAL_DATA,
        'segment_assumptions': SEGMENT_ASSUMPTIONS,
        'scenarios': scenarios,
        'checks': checks,
        'sources': sources,
        'language': 'zh-CN',
        'currency_symbol': '$',
        'unit_suffix': 'B',
        'model_type': 'segment_prediction',  # 标记为分部预测模型
        'fiscal_year_analysis': {
            'fy_end_month': 12,
            'fiscal_year_end': '12-31',
            'is_calendar_year': True,
            'mapping_rule': 'FY ends in December, aligned with calendar year',
            'fiscal_year_note': 'Google fiscal year ends December 31, aligned with calendar year. No mapping needed.',
        },
    }

    out_path = './out/googl_model_data.json'
    import os
    os.makedirs('./out', exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"\n数据已保存到 {out_path}")


if __name__ == '__main__':
    main()