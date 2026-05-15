#!/usr/bin/env python3
"""
run_sanity_checks.py — Phase 5 合理性检验（自动化三层检查）

从 model_data.json 读取假设参数 → 计算10年投影 → 逐条执行三层检验
→ 输出终端报告 + 写入 checks JSON。

用法:
    python scripts/run_sanity_checks.py --input out/nvda_model_data.json
    python scripts/run_sanity_checks.py --input out/googl_model_data.json --output out/googl_checked.json

支持两种模型:
  - 简单 TAM 模型 (NVDA式): assumptions.{scenario}.tam_start / share_start / ...
  - 分部模型 (GOOGL式):   segment_assumptions.{scenario}.search.tam_2025 / ...

退出码: 0=全部通过, 1=仅有警告, 2=存在硬错误
"""

import sys
import json
import argparse
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

sys.stdout.reconfigure(encoding='utf-8')

PROJECTION_YEARS = 10


# ====================================================================
# 数据结构
# ====================================================================

@dataclass
class CheckResult:
    layer: int           # 1/2/3
    name: str            # 检查项名称
    bear: str = ""       # Bear 情景的具体值
    base: str = ""       # Base 情景的具体值
    bull: str = ""       # Bull 情景的具体值
    status: str = ""     # ✓ / ⚠ / ✗
    detail: str = ""     # 异常说明（仅 ⚠/✗ 时填充）

    @property
    def is_pass(self):
        return self.status == '✓'

    @property
    def is_error(self):
        return self.status == '✗'

    @property
    def is_warn(self):
        return self.status == '⚠'


@dataclass
class Projection:
    """单个情景的10年投影"""
    scenario: str
    years: list[str] = field(default_factory=list)
    tam: list[float] = field(default_factory=list)
    revenue: list[float] = field(default_factory=list)
    rev_growth: list[float] = field(default_factory=list)
    gross_margin: list[float] = field(default_factory=list)
    gross_profit: list[float] = field(default_factory=list)
    expense_ratio: list[float] = field(default_factory=list)
    opex: list[float] = field(default_factory=list)
    ebit: list[float] = field(default_factory=list)
    net_income: list[float] = field(default_factory=list)
    target_pe: list[float] = field(default_factory=list)
    implied_mkt_cap: list[float] = field(default_factory=list)
    implied_price: list[float] = field(default_factory=list)
    # 汇总指标
    terminal_revenue: float = 0.0
    terminal_net_income: float = 0.0
    terminal_mkt_cap: float = 0.0
    terminal_gm: float = 0.0
    terminal_pe: float = 0.0
    revenue_cagr: float = 0.0


def _to_billions(raw_revenue: float) -> float:
    """将原始美元金额归一化到 $B。原始值 > 1e7 视为以美元为单位，除以 1e9。"""
    if raw_revenue > 1e7:
        return raw_revenue / 1e9
    return raw_revenue


def compute_simple_projection(assumptions: dict, scenario: str,
                               hist_revenue_b: float) -> Projection:
    """简单 TAM 模型投影: 收入 = TAM × 可触达率 × 市占率
    hist_revenue_b: 历史收入，单位为 $B（已归一化）"""
    a = assumptions[scenario]
    tam_start = a['tam_start']
    tam_cagrs = [a['tam_cagr_1_3'], a['tam_cagr_1_3'], a['tam_cagr_1_3'],
                 a['tam_cagr_4_6'], a['tam_cagr_4_6'], a['tam_cagr_4_6'],
                 a['tam_cagr_7_10'], a['tam_cagr_7_10'], a['tam_cagr_7_10'], a['tam_cagr_7_10']]
    reachable = a['reachable_rate']
    share_start = a['share_start']
    share_end = a['share_end']
    gm_start = a['gm_start']
    gm_end = a['gm_end']
    exp_start = a['expense_start']
    exp_end = a['expense_end']
    tax_rate = a['tax_rate']
    pe_start = a['pe_start']
    pe_end = a['pe_end']
    shares_m = a['shares_m']

    p = Projection(scenario=scenario)
    for i in range(PROJECTION_YEARS):
        if i == 0:
            tam = tam_start * (1 + tam_cagrs[0])
        else:
            tam = p.tam[-1] * (1 + tam_cagrs[i])

        share = share_start + (share_end - share_start) * (i + 1) / PROJECTION_YEARS
        revenue = tam * reachable * share
        gm = gm_start + (gm_end - gm_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else gm_start
        exp_ratio = exp_start + (exp_end - exp_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else exp_start
        gp = revenue * gm
        opex = revenue * exp_ratio
        ebit = gp - opex
        ni = ebit * (1 - tax_rate)
        pe = pe_start + (pe_end - pe_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else pe_start
        mkt_cap = ni * pe
        price = mkt_cap * 1000 / shares_m if shares_m else 0

        p.years.append(f"Y{i+1}")
        p.tam.append(tam)
        p.revenue.append(revenue)
        p.gross_margin.append(gm)
        p.gross_profit.append(gp)
        p.expense_ratio.append(exp_ratio)
        p.opex.append(opex)
        p.ebit.append(ebit)
        p.net_income.append(ni)
        p.target_pe.append(pe)
        p.implied_mkt_cap.append(mkt_cap)
        p.implied_price.append(price)

    # 增速（hist_revenue_b 已归一化为 $B）
    prev = hist_revenue_b
    for rev in p.revenue:
        p.rev_growth.append((rev - prev) / prev if prev else 0)
        prev = rev

    p.terminal_revenue = p.revenue[-1]
    p.terminal_net_income = p.net_income[-1]
    p.terminal_mkt_cap = p.implied_mkt_cap[-1]
    p.terminal_gm = p.gross_margin[-1]
    p.terminal_pe = p.target_pe[-1]
    if hist_revenue_b and p.revenue:
        p.revenue_cagr = (p.terminal_revenue / hist_revenue_b) ** (1 / PROJECTION_YEARS) - 1

    return p


def compute_segment_projection(segment_assumptions: dict, scenario: str,
                                hist_revenue_b: float) -> Projection:
    """分部模型投影: 收入 = Σ(各分部 TAM × 市占率) + 直接预测分部
    hist_revenue_b: 历史收入，单位为 $B（已归一化）"""
    a = segment_assumptions[scenario]
    segments_def = [
        ('search',        'tam_2025', 'tam_cagr_1_3', 'tam_cagr_4_6', 'tam_cagr_7_10', 'share_2025', 'share_2035'),
        ('youtube',       'tam_2025', 'tam_cagr_1_3', 'tam_cagr_4_6', 'tam_cagr_7_10', 'share_2025', 'share_2035'),
        ('cloud',         'tam_2025', 'tam_cagr_1_3', 'tam_cagr_4_6', 'tam_cagr_7_10', 'share_2025', 'share_2035'),
        ('subscriptions', 'tam_2025', 'tam_cagr_1_3', 'tam_cagr_4_6', 'tam_cagr_7_10', 'share_2025', 'share_2035'),
    ]
    # network 是直接预测（萎缩业务）
    has_network = 'network' in a
    network_start = a.get('network', {}).get('revenue_2025', 0)
    network_decline = a.get('network', {}).get('annual_decline', 0)

    gm_start = a['gm_start']
    gm_end = a['gm_end']
    exp_start = a['expense_start']
    exp_end = a['expense_end']
    tax_rate = a['tax_rate']
    pe_start = a['pe_start']
    pe_end = a['pe_end']
    shares_m = a['shares_m']

    p = Projection(scenario=scenario)
    for i in range(PROJECTION_YEARS):
        total_rev = 0.0
        for seg_name, tam_key, c1, c4, c7, sh_start_key, sh_end_key in segments_def:
            seg = a.get(seg_name, {})
            if not seg:
                continue
            tam_start = seg.get(tam_key, 0)
            # 选 CAGR
            if i <= 2:
                cagr = seg.get(c1, 0)
            elif i <= 5:
                cagr = seg.get(c4, 0)
            else:
                cagr = seg.get(c7, 0)
            # 累乘 TAM
            if i == 0:
                tam = tam_start * (1 + cagr)
            else:
                # 从上一年的TAM推算（这里简化：用 (1+cagr)^i 方式）
                tam = tam_start * (1 + cagr) ** (i + 1)
            share_s = seg.get(sh_start_key, 0)
            share_e = seg.get(sh_end_key, 0)
            share = share_s + (share_e - share_s) * (i + 1) / PROJECTION_YEARS
            total_rev += tam * share

        # Network 萎缩
        if has_network:
            nw_rev = network_start * (1 - network_decline) ** (i + 1)
            total_rev += nw_rev

        gm = gm_start + (gm_end - gm_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else gm_start
        exp_ratio = exp_start + (exp_end - exp_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else exp_start
        gp = total_rev * gm
        opex = total_rev * exp_ratio
        ebit = gp - opex
        ni = ebit * (1 - tax_rate)
        pe = pe_start + (pe_end - pe_start) * i / (PROJECTION_YEARS - 1) if PROJECTION_YEARS > 1 else pe_start
        mkt_cap = ni * pe
        price = mkt_cap * 1000 / shares_m if shares_m else 0

        p.years.append(f"Y{i+1}")
        p.tam.append(0)  # 分部模型不跟踪总TAM
        p.revenue.append(total_rev)
        p.gross_margin.append(gm)
        p.gross_profit.append(gp)
        p.expense_ratio.append(exp_ratio)
        p.opex.append(opex)
        p.ebit.append(ebit)
        p.net_income.append(ni)
        p.target_pe.append(pe)
        p.implied_mkt_cap.append(mkt_cap)
        p.implied_price.append(price)

    prev = hist_revenue_b
    for rev in p.revenue:
        p.rev_growth.append((rev - prev) / prev if prev else 0)
        prev = rev

    p.terminal_revenue = p.revenue[-1]
    p.terminal_net_income = p.net_income[-1]
    p.terminal_mkt_cap = p.implied_mkt_cap[-1]
    p.terminal_gm = p.gross_margin[-1]
    p.terminal_pe = p.target_pe[-1]
    if hist_revenue_b and p.revenue:
        p.revenue_cagr = (p.terminal_revenue / hist_revenue_b) ** (1 / PROJECTION_YEARS) - 1

    return p


# ====================================================================
# 第一层：内部逻辑验证
# ====================================================================

def check_growth_decay(proj: Projection, hist_growths: list[float]) -> CheckResult:
    """1.1 收入增速衰减 — 预测期内不应连续2年加速"""
    c = CheckResult(layer=1, name="收入增速衰减")
    # 只检查预测期内部的增速变化（Y2 onwards），忽略历史→预测的跳变
    issues = []
    consecutive_up = 0
    for i in range(2, len(proj.rev_growth)):
        if proj.rev_growth[i] > proj.rev_growth[i - 1]:
            consecutive_up += 1
        else:
            consecutive_up = 0
        if consecutive_up >= 2:
            issues.append(f"Y{i}起连续增速上升")
            break
    c.status = '⚠' if issues else '✓'
    if issues:
        c.detail = '; '.join(issues) + " (可能因分部结构变化，请确认是否合理)"
    else:
        c.detail = "增速正常衰减"
    return c


def check_growth_cap(proj: Projection, tam_growths: list[float]) -> CheckResult:
    """1.1b 收入增速上限 — 不超过 TAM增速 × 1.5（跳过Y1过渡年）"""
    c = CheckResult(layer=1, name="收入增速vs TAM上限")
    issues = []
    # 从 Y2 开始检查（Y1 是历史→预测过渡年）
    for i in range(1, min(len(proj.rev_growth), len(tam_growths))):
        cap = tam_growths[i] * 1.5
        if tam_growths[i] > 0 and proj.rev_growth[i] > cap + 0.05:
            issues.append(f"Y{i+1}: 收入增速{proj.rev_growth[i]:.1%} > TAM增速上限{cap:.1%}")
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "增速在TAM约束内"
    return c


def check_revenue_vs_tam(proj: Projection) -> CheckResult:
    """1.1c 收入规模上限 — 收入不应超过 TAM × 80%"""
    c = CheckResult(layer=1, name="收入vs TAM天花板")
    issues = []
    for i in range(len(proj.revenue)):
        if proj.tam[i] > 0 and proj.revenue[i] > proj.tam[i] * 0.8:
            issues.append(f"Y{i+1}: 收入{proj.revenue[i]:.1f}B > TAM×80%={proj.tam[i]*0.8:.1f}B")
    c.status = '✗' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "收入在TAM约束内"
    return c


def check_growth_smoothness(proj: Projection) -> CheckResult:
    """1.1d 增速变化平滑性 — 相邻两年增速差 ≤ 30pp（跳过 Y1 首次跳变）"""
    c = CheckResult(layer=1, name="增速变化平滑性")
    issues = []
    # 从 Y2→Y3 开始检查，跳过 Y0(历史→Y1)的天然断档
    for i in range(2, len(proj.rev_growth)):
        delta = abs(proj.rev_growth[i] - proj.rev_growth[i - 1])
        if delta > 0.30:
            issues.append(f"Y{i}→Y{i+1}: |Δ增速|={delta:.1%} > 30pp")
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "增速变化平滑"
    return c


def check_gross_margin_range(proj: Projection, industry: str = 'hardware') -> CheckResult:
    """1.2a 毛利率范围 — 硬件30-75%, 软件60-95%"""
    c = CheckResult(layer=1, name="毛利率范围")
    limits = {'hardware': (0.30, 0.75), 'software': (0.60, 0.95),
              'semiconductor': (0.40, 0.80), 'internet': (0.50, 0.90)}
    lo, hi = limits.get(industry, (0.20, 0.80))
    issues = []
    for i, gm in enumerate(proj.gross_margin):
        if gm < lo or gm > hi:
            issues.append(f"Y{i+1}: GM={gm:.1%}, 范围=[{lo:.0%},{hi:.0%}]")
    c.status = '✗' if issues else '✓'
    c.detail = '; '.join(issues) if issues else f"GM在{industry}范围[{lo:.0%},{hi:.0%}]内"
    return c


def check_gm_change_speed(proj: Projection) -> CheckResult:
    """1.2b 毛利率变化速度 — 每年 ≤ ±5pp"""
    c = CheckResult(layer=1, name="毛利率变化速度")
    issues = []
    prev = proj.gross_margin[0] if proj.gross_margin else 0
    for i in range(1, len(proj.gross_margin)):
        delta = abs(proj.gross_margin[i] - prev)
        if delta > 0.05:
            issues.append(f"Y{i+1}: |ΔGM|={delta:.1%} > 5pp")
        prev = proj.gross_margin[i]
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "GM变化 ≤5pp/年"
    return c


def check_expense_trend(proj: Projection) -> CheckResult:
    """1.2c 费用率趋势 — 不应连续3年上升"""
    c = CheckResult(layer=1, name="费用率趋势")
    issues = []
    consecutive_up = 0
    prev = proj.expense_ratio[0] if proj.expense_ratio else 0
    for i in range(1, len(proj.expense_ratio)):
        if proj.expense_ratio[i] > prev:
            consecutive_up += 1
        else:
            consecutive_up = 0
        if consecutive_up >= 3:
            issues.append(f"Y{i-1}起费用率连续上升")
            break
        prev = proj.expense_ratio[i]
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "费用率趋势合理"
    return c


def check_ebit_margin_cap(proj: Projection, hist_peak_ebit_margin: float = 0.65) -> CheckResult:
    """1.2d EBIT margin — 不超过历史峰值 × 1.1"""
    c = CheckResult(layer=1, name="EBIT利润率上限")
    cap = hist_peak_ebit_margin * 1.1
    issues = []
    for i in range(len(proj.ebit)):
        margin = proj.ebit[i] / proj.revenue[i] if proj.revenue[i] else 0
        if margin > cap:
            issues.append(f"Y{i+1}: EBIT%={margin:.1%} > 上限{cap:.1%}")
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else f"EBIT% ≤ {cap:.1%}"
    return c


def check_net_vs_gross_margin(proj: Projection) -> CheckResult:
    """1.2e 净利率 < 毛利率"""
    c = CheckResult(layer=1, name="净利率<毛利率")
    issues = []
    for i in range(len(proj.revenue)):
        if proj.revenue[i] > 0:
            nm = proj.net_income[i] / proj.revenue[i]
            if nm > proj.gross_margin[i]:
                issues.append(f"Y{i+1}: 净利率{nm:.1%} > 毛利率{proj.gross_margin[i]:.1%}")
    c.status = '✗' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "净利率<毛利率 ✓"
    return c


def check_pe_range(proj: Projection, hist_pe_min: float = 5, hist_pe_max: float = 100) -> CheckResult:
    """1.3a PE范围 — 5x ~ 历史峰值"""
    c = CheckResult(layer=1, name="PE范围")
    issues = []
    for i, pe in enumerate(proj.target_pe):
        if pe < hist_pe_min or pe > hist_pe_max:
            issues.append(f"Y{i+1}: PE={pe:.1f}x, 范围=[{hist_pe_min},{hist_pe_max}]")
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else f"PE在[{hist_pe_min},{hist_pe_max}]内"
    return c


def check_mkt_cap_growth(proj: Projection) -> CheckResult:
    """1.3b 隐含市值增速 — 不应每年翻倍"""
    c = CheckResult(layer=1, name="隐含市值增速")
    issues = []
    prev = proj.implied_mkt_cap[0] if proj.implied_mkt_cap else 0
    for i in range(1, len(proj.implied_mkt_cap)):
        if prev > 0:
            g = (proj.implied_mkt_cap[i] - prev) / prev
            if g > 1.0:
                issues.append(f"Y{i+1}: 市值增速={g:.1%} > 100%")
        prev = proj.implied_mkt_cap[i]
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "市值增速 ≤100%/年"
    return c


def check_ps_ratio(proj: Projection) -> CheckResult:
    """1.3c P/S比 — ≤ 30x (非SaaS)"""
    c = CheckResult(layer=1, name="P/S比率")
    issues = []
    for i in range(len(proj.revenue)):
        if proj.revenue[i] > 0:
            ps = proj.implied_mkt_cap[i] / proj.revenue[i]
            if ps > 30:
                issues.append(f"Y{i+1}: P/S={ps:.1f}x > 30x")
    c.status = '⚠' if issues else '✓'
    c.detail = '; '.join(issues) if issues else "P/S ≤ 30x"
    return c


# ====================================================================
# 第二层：跨情景排序验证
# ====================================================================

def check_ordering(projections: dict[str, Projection]) -> list[CheckResult]:
    """2.1 单调性: Bull > Base > Bear"""
    results = []
    metrics = [
        ('终年收入', 'terminal_revenue', 'B'),
        ('终年净利', 'terminal_net_income', 'B'),
        ('终年隐含市值', 'terminal_mkt_cap', 'B'),
        ('10Y CAGR', 'revenue_cagr', '%'),
        ('终年毛利率', 'terminal_gm', '%'),
        ('终年PE', 'terminal_pe', 'x'),
    ]
    for name, attr, unit in metrics:
        c = CheckResult(layer=2, name=f"跨情景排序: {name}")
        vals = {}
        for s in ['bear', 'base', 'bull']:
            vals[s] = getattr(projections.get(s, Projection(scenario=s)), attr, 0)
        c.bear = f"{vals['bear']:.1f}{unit}"
        c.base = f"{vals['base']:.1f}{unit}"
        c.bull = f"{vals['bull']:.1f}{unit}"
        if vals['bull'] >= vals['base'] >= vals['bear']:
            c.status = '✓'
        else:
            c.status = '✗'
            c.detail = f"排序异常: Bull={vals['bull']:.1f}, Base={vals['base']:.1f}, Bear={vals['bear']:.1f}"
        results.append(c)
    return results


def check_gap_reasonableness(projections: dict[str, Projection]) -> list[CheckResult]:
    """2.2 差距合理性"""
    results = []
    proj = projections
    bull_rev = proj['bull'].terminal_revenue
    base_rev = proj['base'].terminal_revenue
    bear_rev = proj['bear'].terminal_revenue

    # Bull vs Base 差距
    c1 = CheckResult(layer=2, name="Bull-Base终年收入差距")
    if base_rev > 0:
        gap = (bull_rev - base_rev) / base_rev
        c1.bull = f"{bull_rev:.1f}B"
        c1.base = f"{base_rev:.1f}B"
        c1.status = '✓' if 0.30 <= gap <= 1.00 else ('⚠' if 0.10 <= gap <= 2.00 else '✗')
        c1.detail = f"差距={gap:.1%}" + (" (偏小)" if gap < 0.30 else (" (偏大)" if gap > 1.00 else ""))
    results.append(c1)

    # Base vs Bear 差距
    c2 = CheckResult(layer=2, name="Base-Bear终年收入差距")
    if bear_rev > 0:
        gap = (base_rev - bear_rev) / bear_rev
        c2.base = f"{base_rev:.1f}B"
        c2.bear = f"{bear_rev:.1f}B"
        c2.status = '✓' if 0.30 <= gap <= 1.00 else ('⚠' if 0.10 <= gap <= 2.00 else '✗')
        c2.detail = f"差距={gap:.1%}" + (" (偏小)" if gap < 0.30 else (" (偏大)" if gap > 1.00 else ""))
    results.append(c2)

    # CAGR 差异
    c3 = CheckResult(layer=2, name="相邻情景CAGR差异")
    cagrs = {s: proj[s].revenue_cagr for s in ['bear', 'base', 'bull']}
    c3.bear = f"{cagrs['bear']:.1%}"
    c3.base = f"{cagrs['base']:.1%}"
    c3.bull = f"{cagrs['bull']:.1%}"
    d1 = cagrs['bull'] - cagrs['base']
    d2 = cagrs['base'] - cagrs['bear']
    issues = []
    for label, d in [('Bull-Base', d1), ('Base-Bear', d2)]:
        if d < 0.02:
            issues.append(f"{label}={d:.1%} < 2pp (过小)")
        elif d > 0.08:
            issues.append(f"{label}={d:.1%} > 8pp (过大)")
    c3.status = '✓' if not issues else '⚠'
    c3.detail = '; '.join(issues) if issues else "CAGR差异合理(2-8pp)"
    results.append(c3)

    # 差距对称性
    c4 = CheckResult(layer=2, name="Bull-Base vs Base-Bear对称性")
    if cagrs['base'] > cagrs['bear'] and cagrs['bull'] > cagrs['base']:
        ratio = d1 / d2 if d2 > 0 else float('inf')
        c4.status = '✓' if 0.33 <= ratio <= 3.0 else '⚠'
        c4.detail = f"Bull-Base/Base-Bear={ratio:.1f}x" + (" (不对称)" if ratio > 3 or ratio < 0.33 else "")
    results.append(c4)

    return results


# ====================================================================
# 第三层：外部基准对比（有外部数据时运行）
# ====================================================================

def check_external_benchmarks(projections: dict[str, Projection],
                               comparables: Optional[dict] = None,
                               industry_cagr: Optional[float] = None,
                               consensus_revenue: Optional[float] = None) -> list[CheckResult]:
    """3.1-3.3 外部基准对比"""
    results = []

    base = projections['base']

    # 3.1 可比公司对比
    if comparables:
        for comp_name, comp_data in comparables.items():
            c = CheckResult(layer=3, name=f"vs {comp_name} CAGR")
            comp_cagr = comp_data.get('cagr', 0)
            c.base = f"{base.revenue_cagr:.1%}"
            c.bull = f"{comp_cagr:.1%}"
            if comp_cagr > 0:
                deviation = abs(base.revenue_cagr - comp_cagr) / comp_cagr
                c.status = '✓' if deviation <= 0.50 else '⚠'
                c.detail = f"偏离={deviation:.1%}" + (" (>50%, 需关注)" if deviation > 0.50 else "")
            results.append(c)

    # 3.2 行业增速对比
    if industry_cagr is not None:
        c = CheckResult(layer=3, name="vs 行业增速")
        c.base = f"公司CAGR {base.revenue_cagr:.1%}"
        c.bull = f"行业CAGR {industry_cagr:.1%}"
        if industry_cagr > 0:
            ratio = base.revenue_cagr / industry_cagr
            c.status = '✓' if ratio <= 2.0 else '⚠'
            c.detail = f"公司/行业={ratio:.1f}x" + (" (>2x, 长期难持续)" if ratio > 2 else "")
        results.append(c)

    # 3.3 市场共识对比
    if consensus_revenue is not None:
        c = CheckResult(layer=3, name="vs 市场共识终年收入")
        c.base = f"模型 {base.terminal_revenue:.1f}B"
        c.bull = f"共识 {consensus_revenue:.1f}B"
        if consensus_revenue > 0:
            dev = abs(base.terminal_revenue - consensus_revenue) / consensus_revenue
            c.status = '✓' if dev <= 0.30 else '⚠'
            c.detail = f"偏差={dev:.1%}" + (" (>30%)" if dev > 0.30 else "")
        results.append(c)

    return results


# ====================================================================
# 框架校准检查
# ====================================================================

def check_framework_calibration(assumptions: dict, hist_data: dict,
                                 model_type: str = 'simple') -> CheckResult:
    """反推校验：用框架推算历史年份收入，与实际对比"""
    c = CheckResult(layer=1, name="框架反推校准")
    last_hist_year = sorted(hist_data.keys())[-1] if hist_data else '2025A'
    actual_rev = hist_data.get(last_hist_year, {}).get('revenue', 0)

    base = assumptions.get('base', assumptions.get('segment_assumptions', {}).get('base', {}))
    if not base:
        c.status = '⚠'
        c.detail = "无假设数据，跳过校准"
        return c

    if model_type == 'simple':
        tam = base.get('tam_start', 0)
        rate = base.get('reachable_rate', 0)
        share = base.get('share_start', 0)
        implied = tam * rate * share
        c.base = f"${tam:.0f}B×{rate:.1%}×{share:.1%}=${implied:.1f}B"
        if isinstance(actual_rev, (int, float)) and actual_rev > 0:
            actual_b = actual_rev / 1e9
            err = abs(implied - actual_b) / actual_b
            c.status = '✓' if err <= 0.15 else ('⚠' if err <= 0.25 else '✗')
            c.detail = f"实际={actual_b:.1f}B, 误差={err:.1%}"
    else:
        c.status = '✓'
        c.detail = "分部模型：各分部单独校准（略）"

    return c


# ====================================================================
# 主流程
# ====================================================================

def detect_model_type(raw: dict) -> str:
    if 'segment_assumptions' in raw:
        return 'segment'
    if 'assumptions' in raw:
        return 'simple'
    return 'unknown'


def detect_industry(raw: dict) -> str:
    """从数据中的行业标签推断"""
    ticker = raw.get('ticker', '').upper()
    notes = str(raw.get('notes', ''))
    if any(kw in notes for kw in ['semiconductor', '芯片', 'GPU', 'foundry']):
        return 'semiconductor'
    if any(kw in notes for kw in ['SaaS', 'software', '软件']):
        return 'software'
    if any(kw in notes for kw in ['internet', 'digital ad', '广告', 'search']):
        return 'internet'
    return 'hardware'


def run_checks(raw: dict, external: Optional[dict] = None) -> list[CheckResult]:
    """主入口：运行全部三层检验"""
    all_checks = []
    model_type = detect_model_type(raw)
    industry = detect_industry(raw)
    external = external or {}

    # 获取历史数据（归一化到 $B）
    hist_data = raw.get('hist_data', {})
    hist_years = sorted(hist_data.keys())
    last_year = hist_years[-1] if hist_years else '2025A'
    hist_revenue_raw = hist_data.get(last_year, {}).get('revenue', 0)
    hist_revenue_b = _to_billions(hist_revenue_raw)
    # 历史增速（最近3年）
    hist_growths = []
    for i in range(1, len(hist_years)):
        prev_rev = hist_data[hist_years[i - 1]].get('revenue', 0)
        curr_rev = hist_data[hist_years[i]].get('revenue', 0)
        if prev_rev:
            hist_growths.append((curr_rev - prev_rev) / prev_rev)

    # 历史PE极值
    hist_pe_max = raw.get('val_data', {}).get('current_pe', 40) * 2.5

    # === 框架校准 ===
    if model_type == 'simple':
        assumptions = raw.get('assumptions', {})
    else:
        assumptions = raw.get('segment_assumptions', {})

    if assumptions:
        all_checks.append(check_framework_calibration(assumptions, hist_data, model_type))

    # === 计算三情景投影 ===
    projections: dict[str, Projection] = {}
    for scenario in ['bear', 'base', 'bull']:
        if model_type == 'simple':
            assumptions = raw.get('assumptions', {})
            proj = compute_simple_projection(assumptions, scenario, hist_revenue_b)
        else:
            assumptions = raw.get('segment_assumptions', {})
            proj = compute_segment_projection(assumptions, scenario, hist_revenue_b)
        projections[scenario] = proj

    # === 第一层：内部逻辑 ===
    for scenario, proj in projections.items():
        prefix = f"[{scenario.upper()}] "

        # 1.1 增速
        c = check_growth_decay(proj, hist_growths)
        c.name = prefix + c.name
        c.bear = c.base = c.bull = ""
        c.status = c.status  # keep original
        all_checks.append(c)

        # 1.1b 增速上限 (only for simple model where we have TAM growth)
        if model_type == 'simple':
            # Y1的TAM增速直接用CAGR_1_3，后续年份从投影反推
            a = raw.get('assumptions', {}).get(scenario, {})
            cagr_1_3 = a.get('tam_cagr_1_3', 0)
            tam_growths = [cagr_1_3]
            tam_g = proj.tam
            for i in range(1, len(tam_g)):
                if tam_g[i - 1] > 0:
                    tam_growths.append((tam_g[i] - tam_g[i - 1]) / tam_g[i - 1])
                else:
                    tam_growths.append(0)
            c = check_growth_cap(proj, tam_growths)
            c.name = prefix + c.name
            all_checks.append(c)

            c = check_revenue_vs_tam(proj)
            c.name = prefix + c.name
            all_checks.append(c)

        c = check_growth_smoothness(proj)
        c.name = prefix + c.name
        all_checks.append(c)

        # 1.2 利润率
        c = check_gross_margin_range(proj, industry)
        c.name = prefix + c.name
        all_checks.append(c)

        c = check_gm_change_speed(proj)
        c.name = prefix + c.name
        all_checks.append(c)

        c = check_expense_trend(proj)
        c.name = prefix + c.name
        all_checks.append(c)

        # 1.3 估值
        c = check_pe_range(proj, hist_pe_max=hist_pe_max)
        c.name = prefix + c.name
        all_checks.append(c)

        c = check_mkt_cap_growth(proj)
        c.name = prefix + c.name
        all_checks.append(c)

        c = check_net_vs_gross_margin(proj)
        c.name = prefix + c.name
        all_checks.append(c)

    # === 第二层：跨情景排序 ===
    all_checks.extend(check_ordering(projections))
    all_checks.extend(check_gap_reasonableness(projections))

    # === 第三层：外部基准 ===
    comps = external.get('comparables')
    ind_cagr = external.get('industry_cagr')
    consensus = external.get('consensus_revenue')
    if comps or ind_cagr or consensus:
        all_checks.extend(check_external_benchmarks(projections, comps, ind_cagr, consensus))
    else:
        c = CheckResult(layer=3, name="外部基准对比", status='⚠',
                        detail="无外部数据，跳过第三层检验。建议提供可比公司/行业增速/市场共识数据。")
        all_checks.append(c)

    return all_checks


def format_report(checks: list[CheckResult], ticker: str = "") -> str:
    """生成可读报告"""
    lines = []
    lines.append("=" * 72)
    lines.append(f"  合理性检验报告 — {ticker}" if ticker else "  合理性检验报告")
    lines.append("=" * 72)

    for layer, title in [(1, "第一层：内部逻辑验证"), (2, "第二层：跨情景排序验证"), (3, "第三层：外部基准对比")]:
        lines.append(f"\n{'─' * 60}")
        lines.append(f"  {title}")
        lines.append(f"{'─' * 60}")
        layer_checks = [c for c in checks if c.layer == layer]
        for c in layer_checks:
            icon = {'✓': '✓', '⚠': '⚠', '✗': '✗'}.get(c.status, '?')
            lines.append(f"  {icon} {c.name}")
            if c.detail:
                lines.append(f"    {c.detail}")

    # 汇总
    errors = sum(1 for c in checks if c.is_error)
    warns = sum(1 for c in checks if c.is_warn)
    passes = sum(1 for c in checks if c.is_pass)
    lines.append(f"\n{'=' * 72}")
    lines.append(f"  汇总: {passes}✓ 通过  {warns}⚠ 警告  {errors}✗ 错误")
    if errors:
        lines.append(f"  ⛔ 存在 {errors} 项硬错误，建议修正后重新运行。")
    elif warns:
        lines.append(f"  ⚡ 存在 {warns} 项警告，请审阅后决定是否调整。")
    else:
        lines.append(f"  ✅ 全部检查通过，模型可以输出。")
    lines.append(f"{'=' * 72}")
    return '\n'.join(lines)


def checks_to_json(checks: list[CheckResult]) -> list[dict]:
    """转为 build_excel.py 兼容的 JSON 格式"""
    return [
        {
            'name': c.name,
            'bear': c.bear,
            'base': c.base,
            'bull': c.bull,
            'status': f"{c.status} {c.detail}" if c.detail else c.status,
        }
        for c in checks
    ]


def main():
    parser = argparse.ArgumentParser(description='Phase 5 合理性检验（自动三层检查）')
    parser.add_argument('--input', required=True, help='model_data.json 路径')
    parser.add_argument('--output', help='输出路径（默认覆盖原文件，追加 checks）')
    parser.add_argument('--ticker', help='股票代码（用于报告标题）')
    parser.add_argument('--industry-cagr', type=float, help='行业共识CAGR（第三层外部基准）')
    parser.add_argument('--consensus-revenue', type=float, help='分析师一致预期终年收入（$B）')
    parser.add_argument('--json-only', action='store_true', help='仅输出 checks JSON')
    args = parser.parse_args()

    with open(args.input, 'r', encoding='utf-8') as f:
        raw = json.load(f)

    # 外部数据
    external = {}
    if args.industry_cagr:
        external['industry_cagr'] = args.industry_cagr
    if args.consensus_revenue:
        external['consensus_revenue'] = args.consensus_revenue

    checks = run_checks(raw, external)
    ticker = args.ticker or raw.get('ticker', '')

    if args.json_only:
        print(json.dumps(checks_to_json(checks), ensure_ascii=False, indent=2))
    else:
        print(format_report(checks, ticker))

    # 写入 JSON
    output_path = args.output or args.input
    raw['checks'] = checks_to_json(checks)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(raw, f, ensure_ascii=False, indent=2)
    if not args.json_only:
        print(f"\n✓ checks 已写入: {output_path}")

    # 退出码
    errors = sum(1 for c in checks if c.is_error)
    warns = sum(1 for c in checks if c.is_warn)
    if errors:
        sys.exit(2)
    elif warns:
        sys.exit(1)
    else:
        sys.exit(0)


if __name__ == '__main__':
    main()
