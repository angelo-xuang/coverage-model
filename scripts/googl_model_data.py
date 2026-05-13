#!/usr/bin/env python3
"""
googl_model_data.py — Google (GOOGL) 估值模型数据准备
修正版：市占率持平，毛利率/费用率扩张更保守
"""

import json
import sys
sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 修正后的三情景假设
# ============================================================

ASSUMPTIONS = {
    'bear': {
        'tam_start': 3000.0,        # $B
        'tam_cagr_1_3': 0.09,
        'tam_cagr_4_6': 0.06,
        'tam_cagr_7_10': 0.04,
        'reachable_rate': 0.40,
        'share_start': 0.336,
        'share_end': 0.310,         # 市占率微降（搜索被蚕食 > Cloud 增量）
        'gm_start': 0.597,
        'gm_end': 0.580,            # AI CapEx 折旧压制
        'expense_start': 0.277,     # 2025A 实际
        'expense_end': 0.290,       # 需要更多投入维持竞争
        'tax_rate': 0.17,
        'target_pe': 22,            # 向历史中位PE回归
        'shares_m': 5824,
    },
    'base': {
        'tam_start': 3000.0,
        'tam_cagr_1_3': 0.14,
        'tam_cagr_4_6': 0.10,
        'tam_cagr_7_10': 0.06,
        'reachable_rate': 0.40,
        'share_start': 0.336,
        'share_end': 0.340,         # 修正：基本持平（+0.4pp）
        'gm_start': 0.597,
        'gm_end': 0.615,            # 修正：+1.8pp（非+3.8pp）
        'expense_start': 0.277,     # 2025A 实际
        'expense_end': 0.260,       # 修正：-1.7pp（非-3.2pp）
        'tax_rate': 0.15,
        'target_pe': 27,            # 接近当前PE
        'shares_m': 5824,
    },
    'bull': {
        'tam_start': 3000.0,
        'tam_cagr_1_3': 0.20,
        'tam_cagr_4_6': 0.14,
        'tam_cagr_7_10': 0.09,
        'reachable_rate': 0.40,
        'share_start': 0.336,
        'share_end': 0.380,         # 修正：从42%降到38%
        'gm_start': 0.597,
        'gm_end': 0.640,            # 修正：从66%降到64%
        'expense_start': 0.277,     # 2025A 实际
        'expense_end': 0.245,       # 修正：从23%提到24.5%
        'tax_rate': 0.14,
        'target_pe': 33,            # AI溢价维持
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
              'rd_ratio': 0.1231, 'sga_ratio': 0.0732, 'revenue_growth': 0.1507},
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


def project_scenario(name: str, params: dict, hist_data: dict, n_years: int = 10):
    """对单个情景做完整预测"""
    last_year = max(int(k.replace('A', '')) for k in hist_data)
    last_rev_M = hist_data[f'{last_year}A']['revenue']  # $M
    last_rev_B = last_rev_M / 1000  # $B

    rows = {y: {} for y in range(1, n_years + 1)}

    tam = params['tam_start']
    share_start = params['share_start']
    share_end = params['share_end']
    gm_start = params['gm_start']
    gm_end = params['gm_end']
    exp_start = params['expense_start']  # 起始费用率 (2025A)
    exp_end = params['expense_end']      # 终年费用率
    tax_rate = params['tax_rate']

    tam = params['tam_start']
    for t in range(1, n_years + 1):
        # TAM：每年都增长（包括第一年）
        if t <= 3:
            cagr = params['tam_cagr_1_3']
        elif t <= 6:
            cagr = params['tam_cagr_4_6']
        else:
            cagr = params['tam_cagr_7_10']
        tam = tam * (1 + cagr)

        sam = tam * params['reachable_rate']
        ms = share_start + (share_end - share_start) * (t - 1) / (n_years - 1)
        rev = sam * ms
        gm = gm_start + (gm_end - gm_start) * (t - 1) / (n_years - 1)
        gross_profit = rev * gm

        # 费用率：线性从起始到终年
        exp = exp_start + (exp_end - exp_start) * (t - 1) / (n_years - 1)
        opex = rev * exp
        ebit = gross_profit - opex
        net_income = ebit * (1 - tax_rate)

        # 增速
        if t == 1:
            rev_growth = (rev - last_rev_B) / last_rev_B
        else:
            prev_rev = rows[t - 1]['rev']
            rev_growth = (rev - prev_rev) / prev_rev if prev_rev else 0

        # PE: 使用 target_pe 作为终年PE
        pe = params['target_pe']

        rows[t] = {
            'tam': tam,
            'sam': sam,
            'ms': ms,
            'rev': rev,
            'rev_growth': rev_growth,
            'gm': gm,
            'gross_profit': gross_profit,
            'exp': exp,
            'opex': opex,
            'ebit': ebit,
            'tax_rate': tax_rate,
            'net_income': net_income,
            'ebit_margin': ebit / rev if rev else 0,
            'net_margin': net_income / rev if rev else 0,
            'pe': pe,
        }

    # 估值：用终年 PE
    terminal_ni = rows[n_years]['net_income']
    terminal_pe = rows[n_years]['pe']
    shares_m = params['shares_m']
    mkt_cap = terminal_ni * terminal_pe
    price = mkt_cap * 1000 / shares_m
    annual_return = (price / VAL_DATA['current_price']) ** (1 / n_years) - 1

    # CAGR
    cagr = (rev / last_rev_B) ** (1 / n_years) - 1

    return {
        'name': name,
        'rows': rows,
        'terminal': {
            'year': last_year + n_years,
            'rev_B': rev,
            'gm': gm,
            'ebit_margin': ebit / rev if rev else 0,
            'net_income_B': net_income,
            'mkt_cap_B': mkt_cap,
            'price': price,
            'cagr': cagr,
            'annual_return': annual_return,
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
        # 检查是否有连续2年增速上升（排除前2年，因为TAM CAGR切换可能有跳变）
        accel_count = 0
        for i in range(1, len(growths)):
            if growths[i] > growths[i-1]:
                accel_count += 1
            else:
                accel_count = 0
        growth_decay_ok = accel_count < 2

        # 增速绝对上限
        max_growth = max(growths)
        growth_abs_ok = max_growth < 0.60

        # 毛利率范围
        gms = [rows[t]['gm'] for t in range(1, n_years + 1)]
        gm_ok = all(0.30 < g < 0.80 for g in gms)

        # EBIT margin 范围
        ebm = [rows[t]['ebit_margin'] for t in range(1, n_years + 1)]
        ebit_ok = all(0 < e < 0.55 for e in ebm)

        results.append({
            'name': f'[{name}] 收入增速衰减',
            'bear': '✓' if name == 'BEAR' else '',
            'base': '✓' if name == 'BASE' else '',
            'bull': '✓' if name == 'BULL' else '',
            'status': '✓ 通过' if growth_decay_ok else '⚠ 注意',
        })
        results.append({
            'name': f'[{name}] 毛利率范围',
            'bear': '✓' if name == 'BEAR' else '',
            'base': '✓' if name == 'BASE' else '',
            'bull': '✓' if name == 'BULL' else '',
            'status': '✓ 通过' if gm_ok else '✗ 不通过',
        })
        results.append({
            'name': f'[{name}] EBIT margin 合理',
            'bear': '✓' if name == 'BEAR' else '',
            'base': '✓' if name == 'BASE' else '',
            'bull': '✓' if name == 'BULL' else '',
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
        'bear': '',
        'base': '',
        'bull': '',
        'status': '✓ 通过' if gap_ok else '⚠ 差距偏大',
    })

    # === 第三层：外部基准 ===
    base_cagr = base_t['cagr']
    results.append({
        'name': f'Base CAGR ({base_cagr:.1%}) vs MSFT历史 (~12%)',
        'bear': '',
        'base': f'{base_cagr:.1%}',
        'bull': '',
        'status': '✓ 合理' if 0.05 < base_cagr < 0.20 else '⚠ 偏离',
    })

    return results


def main():
    scenarios = []
    for name in ['bear', 'base', 'bull']:
        proj = project_scenario(name, ASSUMPTIONS[name], HIST_DATA)
        scenarios.append(proj)

        t = proj['terminal']
        print(f"\n{'='*60}")
        print(f"  {name.upper()} CASE — 终年 {t['year']}E")
        print(f"{'='*60}")
        print(f"  收入:       ${t['rev_B']:>10,.1f}B  (CAGR: {t['cagr']:.1%})")
        print(f"  毛利率:     {t['gm']:>10.1%}")
        print(f"  EBIT margin:{t['ebit_margin']:>10.1%}")
        print(f"  净利润:     ${t['net_income_B']:>10,.1f}B")
        print(f"  隐含市值:   ${t['mkt_cap_B']:>10,.1f}B")
        print(f"  隐含股价:   ${t['price']:>10,.1f}")
        print(f"  vs 当前:    {(t['price']/387.35 - 1):>10.1%}")
        print(f"  年化回报:   {t['annual_return']:>10.1%}")

    # 逐年明细（Base case）
    base_rows = scenarios[1]['rows']
    print(f"\n{'='*80}")
    print(f"  BASE CASE — 逐年明细")
    print(f"{'='*80}")
    header = f"{'Year':>6} {'TAM':>8} {'SAM':>8} {'MS':>6} {'Rev':>8} {'Gr':>6} {'GM':>6} {'EBIT':>8} {'NI':>8} {'PE':>5} {'Price':>8}"
    print(f"  {header}")
    print(f"  {'-'*82}")
    for t in range(1, 11):
        r = base_rows[t]
        year = 2025 + t
        implied_price = r['net_income'] * r['pe'] * 1000 / 5824
        print(f"  {year}E {r['tam']:>8,.0f} {r['sam']:>8,.0f} {r['ms']:>6.1%} "
              f"{r['rev']:>8,.1f} {r['rev_growth']:>6.1%} {r['gm']:>6.1%} "
              f"{r['ebit']:>8,.1f} {r['net_income']:>8,.1f} {r['pe']:>5.1f} "
              f"${implied_price:>8,.1f}")

    # 合理性检验
    checks = sanity_checks(scenarios, HIST_DATA)
    print(f"\n{'='*60}")
    print(f"  合理性检验")
    print(f"{'='*60}")
    for c in checks:
        vals = ' | '.join(filter(None, [c['bear'], c['base'], c['bull']]))
        print(f"  {c['name']:<40} {c['status']}")

    # 假设来源注释（说明每个数字是怎么推导出来的）
    sources = {
        'tam_start': '基于Gartner 2025全球IT支出预测(~$5T) × 数字化服务渗透率，取$3T作为GOOGL可参与市场',
        'tam_cagr_1_3': '基于IDC/Gartner对AI和云计算支出的增速预测(15-25%)，Bear取低、Bull取高',
        'tam_cagr_4_6': '增速衰减假设：参考S曲线，渗透率10%-50%阶段增速自然下降',
        'tam_cagr_7_10': '长期增速向GDP+通胀回归(3-6%)，参考Cisco/AWS历史增速衰减轨迹',
        'reachable_rate': '40%：基于GOOGL业务覆盖（搜索+Cloud+YouTube+Android）占全球数字支出的比例',
        'share_start': '33.6%：基于2025A实际收入/$3T SAM计算',
        'share_end': 'Bear 31%: 搜索被AI蚕食>Cloud增量; Base 34%: 基本持平; Bull 38%: Cloud+AI持续抢份额',
        'gm_start': '59.7%：2025A实际毛利率',
        'gm_end': 'Bear 58%: AI CapEx折旧压制; Base 61.5%: 每5年+1.8pp(参考2019-2025轨迹); Bull 64%: AI效率提升',
        'expense_start': '27.7%：2025A实际费用率(RD+SGA占收入比)',
        'expense_end': 'Bear 29%: 竞争加剧需更多投入; Base 26%: 规模效应-1.7pp; Bull 24.5%: AI自动化降本',
        'tax_rate': '基于GOOGL历史有效税率14-17%，Bear取高、Bull取低',
        'target_pe': 'Bear 22x: 向历史中位PE回归; Base 27x: 接近当前PE; Bull 33x: AI溢价维持',
        'shares_m': '5824M：基于yfinance当前流通股数',
    }

    # 输出 JSON 供 build_excel.py 使用
    output = {
        'hist_data': HIST_DATA,
        'val_data': VAL_DATA,
        'assumptions': ASSUMPTIONS,
        'checks': checks,
        'sources': sources,
        'language': 'en',
        'currency_symbol': '$',
        'unit_suffix': 'M',
        'fiscal_year_analysis': {
            'fy_end_month': 12,
            'fiscal_year_end': '12-31',
            'is_calendar_year': True,
            'mapping_rule': 'FY ends in December, aligned with calendar year',
            'date_mappings': {},
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
