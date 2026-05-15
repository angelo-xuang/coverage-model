#!/usr/bin/env python3
"""
patch_build_excel.py — 对 build_excel.py 做三大改造:
1. build_assumptions_sheet 返回 row_map
2. _build_tam_standard_projections 重写为紧凑布局
3. _build_profit_tail 重写为紧凑布局
4. build_projections 接收 asm_rows
5. main() 串联
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')

FILE = r"C:\Users\Lenovo\.claude\skills\coverage-model\scripts\build_excel.py"

with open(FILE, 'r', encoding='utf-8') as f:
    src = f.read()

# ===== Patch 1: build_assumptions_sheet 返回 row_map =====
# 在函数最后的 ws.column_dimensions 行之后加 return
old_asm_end = """    ws.column_dimensions['G'].width = 55
"""
new_asm_end = """    ws.column_dimensions['G'].width = 55

    # 返回行号映射: {param_key: row_number}
    row_map = {}
    for _ri, _item in enumerate(layout, start=3):
        if _item.get('type') == 'param' and _item.get('key'):
            row_map[_item['key']] = _ri
    return row_map
"""
src = src.replace(old_asm_end, new_asm_end, 1)

# ===== Patch 2: build_projections 接收 asm_rows =====
old_bp = """def build_projections(ws, model_type, assumptions, segments, company_assumptions,
                       hist_data, labels, ccy_sym='$', unit_suffix='B',
                       projection_years=10, current_price=None, extra_segments=None):"""
new_bp = """def build_projections(ws, model_type, assumptions, segments, company_assumptions,
                       hist_data, labels, ccy_sym='$', unit_suffix='B',
                       projection_years=10, current_price=None, extra_segments=None, asm_rows=None):"""
src = src.replace(old_bp, new_bp, 1)

# 更新 build_projections 内部调用，传递 asm_rows
old_bp_call1 = """        _build_tam_standard_projections(ws, assumptions, hist_data, labels, ccy_sym, unit_suffix,
                                         projection_years, current_price, extra_segments)"""
new_bp_call1 = """        _build_tam_standard_projections(ws, assumptions, hist_data, labels, ccy_sym, unit_suffix,
                                         projection_years, current_price, extra_segments, asm_rows)"""
src = src.replace(old_bp_call1, new_bp_call1, 1)


# ===== Patch 3: 重写 _build_tam_standard_projections =====
# 找到函数开始和结束位置
import re

old_tam_func_start = """# ===== TAM Standard 预测 =====
def _build_tam_standard_projections(ws, assumptions, hist_data, labels, ccy_sym, unit_suffix,
                                      N, current_price, extra_segments=None):"""

# 找到下一个函数的开始位置作为结束标记
old_tam_func_end_marker = "\n# ===== 分部 预测 (通用) ====="

idx_start = src.index(old_tam_func_start)
idx_end = src.index(old_tam_func_end_marker)

new_tam_func = r'''# ===== TAM Standard 预测（紧凑布局） =====
def _build_tam_standard_projections(ws, assumptions, hist_data, labels, ccy_sym, unit_suffix,
                                      N, current_price, extra_segments=None, asm_rows=None):
    """紧凑预测：一行一指标，公式直接引用假设表行号"""
    N, _, _, _, _, TC, hist = _forecast_setup(ws, hist_data, labels, N)
    sn = labels['sheet_assumptions']
    h_rev = hist.get('revenue', 0)
    h_rev_B = h_rev / 1000 if h_rev > 1e6 else h_rev
    h_gm = hist.get('gross_margin', 0)
    h_gr = hist.get('revenue_growth', 0)
    h_ni = hist.get('net_income', 0)
    h_ni_B = h_ni / 1000 if h_ni > 1e6 else h_ni

    def a(key):
        if asm_rows and key in asm_rows:
            return f"'{sn}'!$F${asm_rows[key]}"
        # Fallback: compute from fixed layout
        _ar = 4
        _AR = {
            'tam_start': _ar, 'tam_cagr_1_3': _ar+1, 'tam_cagr_4_6': _ar+2, 'tam_cagr_7_10': _ar+3,
            'reachable_rate': _ar+4, 'share_start': _ar+5, 'share_end': _ar+6,
        }
        _ar2 = _ar + 7
        if extra_segments:
            for seg in extra_segments:
                _ar2 += 1 + 4  # subsection + 4 params
        _ar2 += 1  # margin section header
        _AR['gm_start'] = _ar2; _AR['gm_end'] = _ar2+1
        _AR['expense_start'] = _ar2+2; _AR['expense_end'] = _ar2+3
        _AR['tax_rate'] = _ar2+4
        _ar2 += 6  # valuation section header
        _AR['pe_start'] = _ar2; _AR['pe_end'] = _ar2+1
        _AR['shares_m'] = _ar2+2
        return f"'{sn}'!$F${_AR[key]}"

    def a_seg(seg_key, param_key):
        sk = f"{seg_key}_{param_key}"
        if asm_rows and sk in asm_rows:
            return f"'{sn}'!$F${asm_rows[sk]}"
        return a(param_key)  # fallback

    # --- 收入推导 ---
    r = 2
    _S(ws, r, labels['section_rev_forecast'], TC); r += 1

    # TAM: 一行，Y0=起始, Y1-Y10 用 CAGR 连乘
    r_tam = r; _L(ws, r, f"{labels['tam']} ({ccy_sym}{unit_suffix})"); r += 1
    _K(ws, r_tam, 2, f"={a('tam_start')}", '#,##0.0')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col); pc = get_column_letter(col - 1)
        cagr_key = 'tam_cagr_1_3' if i <= 3 else ('tam_cagr_4_6' if i <= 6 else 'tam_cagr_7_10')
        _K(ws, r_tam, col, f"={pc}{r_tam}*(1+{a(cagr_key)})", '#,##0.0')

    # 可触达率: 一行, 常量
    r_reach = r; _L(ws, r, labels['reachable']); r += 1
    for i in range(0, N + 1):
        col = i + 2
        ws.cell(row=r_reach, column=col, value=f"={a('reachable_rate')}")
        style_data_cell(ws, r_reach, col, font=BLACK_FONT, num_fmt='0.0%')

    # 可触达市场: TAM × 可触达率
    r_sam = r; _L(ws, r, f"{labels['sam']} ({ccy_sym}{unit_suffix})"); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_sam, col, f"={c}{r_tam}*{c}{r_reach}", '#,##0.0')

    # 市占率: 一行, 线性插值 (Y0=start, Y10=end)
    r_share = r; _L(ws, r, labels['market_share']); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        if i == 0:
            _K(ws, r_share, col, f"={a('share_start')}", '0.000%')
        else:
            _K(ws, r_share, col,
               f"={a('share_start')}+({a('share_end')}-{a('share_start')})*{i}/{N}", '0.000%')

    # DC 收入: 可触达市场 × 市占率
    r_dc = r; _L(ws, r, f"DC {labels['fc_revenue']} ({ccy_sym}{unit_suffix})"); r += 1
    _B(ws, r_dc, 2, h_rev_B, '#,##0.0')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        _O(ws, r_dc, col, f"={c}{r_sam}*{c}{r_share}")

    # 额外分部
    seg_rev_rows = {'dc': r_dc}
    if extra_segments:
        for seg in extra_segments:
            sk = seg['key']
            _S(ws, r, seg.get('label', sk), TC); r += 1
            r_seg_rev = r
            _L(ws, r, f"{seg.get('label', sk)} {labels['seg_rev']} ({ccy_sym}{unit_suffix})"); r += 1
            for i in range(1, N + 1):
                col = i + 2; c = get_column_letter(col); pc = get_column_letter(col - 1)
                cagr_key = 'cagr_1_3' if i <= 3 else ('cagr_4_6' if i <= 6 else 'cagr_7_10')
                if i == 1:
                    _K(ws, r_seg_rev, col,
                       f"={a_seg(sk, 'rev')}*(1+{a_seg(sk, cagr_key)})", '#,##0.0')
                else:
                    _K(ws, r_seg_rev, col,
                       f"={pc}{r_seg_rev}*(1+{a_seg(sk, cagr_key)})", '#,##0.0')
            seg_rev_rows[sk] = r_seg_rev

    # --- 总收入 ---
    _S(ws, r, labels['total_revenue'], TC); r += 1
    r_total = r; _L(ws, r, f"{labels['total_revenue']} ({ccy_sym}{unit_suffix})"); r += 1
    _B(ws, r_total, 2, h_rev_B, '#,##0.0')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        parts = '+'.join(f"{c}{seg_rev_rows[s]}" for s in seg_rev_rows)
        _O(ws, r_total, col, f"={parts}")

    r_growth = r; _L(ws, r, labels['total_rev_growth']); r += 1
    _B(ws, r_growth, 2, h_gr, '0.0%')
    if h_rev_B:
        _K(ws, r_growth, 3, f"=(C{r_total}-B{r_total})/B{r_total}", '0.0%')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col); pc = get_column_letter(col - 1)
        _K(ws, r_growth, col, f"=({c}{r_total}-{pc}{r_total})/{pc}{r_total}", '0.0%')

    # --- 利润 + 估值 (紧凑) ---
    r = _build_profit_tail_compact(ws, r, asm_rows, labels, ccy_sym, unit_suffix, N, TC,
                                    r_total, hist, h_rev_B, h_gm, h_ni_B, current_price)

    ws.column_dimensions['A'].width = 28
    for c in range(2, TC + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14

'''

src = src[:idx_start] + new_tam_func + src[idx_end:]


# ===== Patch 4: 添加 _build_profit_tail_compact =====
# 在旧 _build_profit_tail 之前插入新函数
old_profit_marker = "def _build_profit_tail(ws, r, asm_co, labels, ccy_sym, unit_suffix, N, TC,"

new_profit_func = r'''def _build_profit_tail_compact(ws, r, asm_rows, labels, ccy_sym, unit_suffix, N, TC,
                          r_total, hist, h_rev_B, h_gm, h_ni_B, current_price):
    """紧凑利润+估值尾部：一行一指标，公式直接引用假设表"""
    VAL = current_price or 100
    sn = labels['sheet_assumptions']

    def a(key):
        if asm_rows and key in asm_rows:
            return f"'{sn}'!$F${asm_rows[key]}"
        # Fallback
        _ar = 4 + 7
        if ws._ws.title == labels.get('sheet_forecast', '预测'):
            pass
        _AR = {
            'gm_start': _ar, 'gm_end': _ar+1,
            'expense_start': _ar+2, 'expense_end': _ar+3,
            'tax_rate': _ar+4,
            'pe_start': _ar+6, 'pe_end': _ar+7,
            'shares_m': _ar+8,
        }
        return f"'{sn}'!$F${_AR[key]}"

    _S(ws, r, labels['section_profit_forecast'], TC); r += 1

    # 毛利率: 一行, Y0=历史, Y1-Y10 从 start 到 end 线性插值
    r_gm = r; _L(ws, r, labels['fc_gm']); r += 1
    _B(ws, r_gm, 2, h_gm, '0.0%')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_gm, col,
           f"={a('gm_start')}+({a('gm_end')}-{a('gm_start')})*{i-1}/{N-1}", '0.0%')

    # 毛利
    r_gp = r; _L(ws, r, f"{labels['fc_gross_profit']} ({ccy_sym}{unit_suffix})"); r += 1
    if h_rev_B and h_gm:
        _B(ws, r_gp, 2, h_rev_B * h_gm, '#,##0.0')
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_gp, col, f"={c}{r_total}*{c}{r_gm}", '#,##0.0')

    # 费用率: 一行, Y0=起始, Y1-Y10 插值
    r_exp = r; _L(ws, r, labels['fc_expense_ratio']); r += 1
    _K(ws, r_exp, 2, f"={a('expense_start')}", '0.0%')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_exp, col,
           f"={a('expense_start')}+({a('expense_end')}-{a('expense_start')})*{i-1}/{N-1}", '0.0%')

    # 费用
    r_opex = r; _L(ws, r, f"{labels['fc_opex']} ({ccy_sym}{unit_suffix})"); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_opex, col, f"={c}{r_total}*{c}{r_exp}", '#,##0.0')

    # EBIT
    r_ebit = r; _L(ws, r, f"{labels['fc_ebit']} ({ccy_sym}{unit_suffix})"); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_ebit, col, f"={c}{r_gp}-{c}{r_opex}", '#,##0.0')

    # 税率: 常量
    r_tax = r; _L(ws, r, labels['fc_tax_rate']); r += 1
    for i in range(0, N + 1):
        col = i + 2
        ws.cell(row=r_tax, column=col, value=f"={a('tax_rate')}")
        style_data_cell(ws, r_tax, col, font=BLACK_FONT, num_fmt='0.0%')

    # 净利润
    r_ni = r; _L(ws, r, f"{labels['fc_net_income']} ({ccy_sym}{unit_suffix})"); r += 1
    _B(ws, r_ni, 2, h_ni_B, '#,##0.0')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        _O(ws, r_ni, col, f"={c}{r_ebit}*(1-{c}{r_tax})", '#,##0.0')

    # 估值
    _S(ws, r, labels['section_val_forecast'], TC); r += 1

    # 目标 PE: 一行, Y0=起始, Y1-Y10 插值
    r_pe = r; _L(ws, r, labels['fc_target_pe']); r += 1
    _K(ws, r_pe, 2, f"={a('pe_start')}", '0.0')
    for i in range(1, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_pe, col,
           f"={a('pe_start')}+({a('pe_end')}-{a('pe_start')})*{i-1}/{N-1}", '0.0')

    # 目标市值
    r_mkt = r; _L(ws, r, f"{labels['fc_implied_mkt_cap']} ({ccy_sym}{unit_suffix})"); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _O(ws, r_mkt, col, f"={c}{r_ni}*{c}{r_pe}", '#,##0.0')

    # 股份数: 常量
    r_sh = r; _L(ws, r, labels['fc_shares']); r += 1
    for i in range(0, N + 1):
        col = i + 2
        ws.cell(row=r_sh, column=col, value=f"={a('shares_m')}")
        style_data_cell(ws, r_sh, col, font=BLACK_FONT, num_fmt='#,##0')

    # 目标股价
    r_price = r; _L(ws, r, f"{labels['implied_price_label']} ({ccy_sym})"); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _O(ws, r_price, col, f"={c}{r_mkt}*1000/{c}{r_sh}", '#,##0.00')

    # 上行空间
    r_prem = r; _L(ws, r, labels['fc_premium']); r += 1
    for i in range(0, N + 1):
        col = i + 2; c = get_column_letter(col)
        _K(ws, r_prem, col, f"=({c}{r_price}-{VAL})/{VAL}", '0.0%')

    return r


'''

src = src.replace(old_profit_marker, new_profit_func + old_profit_marker, 1)


# ===== Patch 5: main() 串联 =====
# 修改 build_assumptions_sheet 调用，接收返回值
old_main_asm = """    build_assumptions_sheet(ws_asm, layout, labels, sources)"""
new_main_asm = """    asm_rows = build_assumptions_sheet(ws_asm, layout, labels, sources)"""
src = src.replace(old_main_asm, new_main_asm, 1)

# 修改 build_projections 调用，传递 asm_rows
old_main_proj = """    build_projections(ws_fc, model_type, assumptions, segments, company_assumptions,
                       data, labels, ccy_sym, asm_unit,
                       args.projection_years, current_price=current_price,
                       extra_segments=extra_segments)"""
new_main_proj = """    build_projections(ws_fc, model_type, assumptions, segments, company_assumptions,
                       data, labels, ccy_sym, asm_unit,
                       args.projection_years, current_price=current_price,
                       extra_segments=extra_segments, asm_rows=asm_rows)"""
src = src.replace(old_main_proj, new_main_proj, 1)


# ===== 写回 =====
with open(FILE, 'w', encoding='utf-8') as f:
    f.write(src)

print("build_excel.py patched successfully!")
print(f"  File: {FILE}")
