#!/usr/bin/env python3
"""
build_excel.py — 生成带场景切换的估值模型 Excel
用法: python build_excel.py --ticker NVDA --input data.json

输入: fetch_financials.py 输出的 JSON 文件（或直接传入参数）
输出: ./out/<ticker>_coverage_model.xlsx

支持:
- 多币种（自动根据公司地区选择 CNY/USD 等）
- 中英文切换（中国运营主体 → 中文，其他 → 英文）
- 假设来源注释列（每个参数旁标明推导依据）
- 精确的 FY→自然年映射说明

依赖: openpyxl
"""

import sys
import json
import argparse
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


# ===== 颜色定义 =====
BLUE_FONT = Font(color="0000FF")
BLACK_FONT = Font(color="000000")
GREEN_FONT = Font(color="008000")
GRAY_FONT = Font(color="808080", italic=True)
WHITE_BOLD = Font(color="FFFFFF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
COL_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
OUTPUT_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
INPUT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
TOGGLE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
NOTE_FILL = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

RIGHT_ALIGN = Alignment(horizontal='right')
WRAP_ALIGN = Alignment(horizontal='left', wrap_text=True, vertical='top')


def style_header_row(ws, row, start_col, end_col, fill=HEADER_FILL, font=WHITE_BOLD):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER


def style_data_cell(ws, row, col, font=BLACK_FONT, fill=None, num_fmt=None):
    cell = ws.cell(row=row, column=col)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.border = THIN_BORDER
    cell.alignment = RIGHT_ALIGN
    if num_fmt:
        cell.number_format = num_fmt


def get_labels(language='en'):
    """获取中英文标签"""
    if language == 'zh-CN':
        return {
            'sheet_hist': '财务历史',
            'sheet_val': '估值历史',
            'sheet_assumptions': '假设',
            'sheet_forecast': '预测',
            'sheet_sanity': '合理性',
            'sheet_fy_info': '财年信息',
            'label_metric': '指标',
            'active_scenario': '活跃情景',
            'active_value': '活跃值',
            'source_label': '来源/推导依据',
            'revenue': '收入',
            'revenue_growth': '收入增速',
            'cogs': '营业成本',
            'gross_profit': '毛利',
            'gross_margin': '毛利率',
            'rd': '研发费用',
            'rd_ratio': '研发费用率',
            'sga': '销售管理费用',
            'sga_ratio': '销售管理费用率',
            'ebit': 'EBIT',
            'ebit_margin': 'EBIT利润率',
            'net_income': '净利润',
            'net_margin': '净利率',
            'current_mkt_cap': '当前市值',
            'current_price': '当前股价',
            'pe_ttm': 'PE (TTM)',
            'pe_fwd': 'PE (Forward)',
            'shares': '流通股数',
            'wk52_high': '52周高',
            'wk52_low': '52周低',
            'section_revenue': '收入驱动',
            'section_margin': '利润率',
            'section_valuation': '估值',
            'section_rev_forecast': '收入推导',
            'section_profit_forecast': '利润推导',
            'section_val_forecast': '估值',
            'tam_start': 'TAM 起始',
            'tam_cagr_13': 'TAM CAGR Y1-Y3',
            'tam_cagr_46': 'TAM CAGR Y4-Y6',
            'tam_cagr_710': 'TAM CAGR Y7-Y10',
            'reachable': '可触达率',
            'share_start': '起始市占率',
            'share_end': '终年市占率',
            'gm_start': '起始毛利率',
            'gm_end': '终年毛利率',
            'exp_start': '起始费用率',
            'exp_end': '终年费用率',
            'tax_rate': '税率',
            'target_pe': '目标 PE',
            'shares_m': '股份数',
            'tam': 'TAM',
            'sam': '可触达市场',
            'market_share': '市占率',
            'fc_revenue': '收入',
            'fc_rev_growth': '收入增速',
            'fc_gm': '毛利率',
            'fc_gross_profit': '毛利',
            'fc_expense_ratio': '费用率',
            'fc_opex': '运营费用',
            'fc_ebit': 'EBIT',
            'fc_tax_rate': '税率',
            'fc_net_income': '净利润',
            'fc_target_pe': '目标 PE',
            'fc_implied_mkt_cap': '隐含市值',
            'fc_shares': '股份数',
            'fc_implied_price': '隐含股价',
            'fc_premium': 'vs 当前股价溢价',
            'implied_price_label': '隐含股价',
        }
    else:
        return {
            'sheet_hist': 'Financial History',
            'sheet_val': 'Valuation',
            'sheet_assumptions': 'Assumptions',
            'sheet_forecast': 'Forecast',
            'sheet_sanity': 'Sanity Checks',
            'sheet_fy_info': 'Fiscal Year Info',
            'label_metric': 'Metric',
            'active_scenario': 'Active Scenario',
            'active_value': 'Active Value',
            'source_label': 'Source / Derivation',
            'revenue': 'Revenue',
            'revenue_growth': 'Rev Growth',
            'cogs': 'COGS',
            'gross_profit': 'Gross Profit',
            'gross_margin': 'Gross Margin',
            'rd': 'R&D',
            'rd_ratio': 'R&D Ratio',
            'sga': 'SG&A',
            'sga_ratio': 'SG&A Ratio',
            'ebit': 'EBIT',
            'ebit_margin': 'EBIT Margin',
            'net_income': 'Net Income',
            'net_margin': 'Net Margin',
            'current_mkt_cap': 'Current Market Cap',
            'current_price': 'Current Price',
            'pe_ttm': 'PE (TTM)',
            'pe_fwd': 'PE (Forward)',
            'shares': 'Shares Outstanding',
            'wk52_high': '52W High',
            'wk52_low': '52W Low',
            'section_revenue': 'Revenue Drivers',
            'section_margin': 'Profitability',
            'section_valuation': 'Valuation',
            'section_rev_forecast': 'Revenue Build-up',
            'section_profit_forecast': 'Profit Build-up',
            'section_val_forecast': 'Valuation',
            'tam_start': 'TAM Starting',
            'tam_cagr_13': 'TAM CAGR Y1-Y3',
            'tam_cagr_46': 'TAM CAGR Y4-Y6',
            'tam_cagr_710': 'TAM CAGR Y7-Y10',
            'reachable': 'Reachable Rate',
            'share_start': 'Starting Share',
            'share_end': 'Terminal Share',
            'gm_start': 'Starting GM',
            'gm_end': 'Terminal GM',
            'exp_start': 'Starting Exp Ratio',
            'exp_end': 'Terminal Exp Ratio',
            'tax_rate': 'Tax Rate',
            'target_pe': 'Target PE',
            'shares_m': 'Shares (M)',
            'tam': 'TAM',
            'sam': 'Addressable Mkt',
            'market_share': 'Market Share',
            'fc_revenue': 'Revenue',
            'fc_rev_growth': 'Rev Growth',
            'fc_gm': 'Gross Margin',
            'fc_gross_profit': 'Gross Profit',
            'fc_expense_ratio': 'Expense Ratio',
            'fc_opex': 'OpEx',
            'fc_ebit': 'EBIT',
            'fc_tax_rate': 'Tax Rate',
            'fc_net_income': 'Net Income',
            'fc_target_pe': 'Target PE',
            'fc_implied_mkt_cap': 'Implied Mkt Cap',
            'fc_shares': 'Shares (M)',
            'fc_implied_price': 'Implied Price',
            'fc_premium': 'vs Current Premium',
            'implied_price_label': 'Implied Price',
        }


def build_financial_history(ws, data, labels, ccy_sym='$', unit_suffix='M'):
    ws.title = labels['sheet_hist']
    years = sorted(data.keys())
    if not years:
        return

    ws.cell(row=1, column=1, value=labels['label_metric'])
    style_header_row(ws, 1, 1, len(years) + 1)
    for i, y in enumerate(years):
        ws.cell(row=1, column=i + 2, value=y)
        style_header_row(ws, 1, i + 2, i + 2, fill=COL_HEADER_FILL, font=BLACK_BOLD)

    rows_def = [
        (f"{labels['revenue']} ({ccy_sym}{unit_suffix})", "revenue", BLUE_FONT, INPUT_FILL, '#,##0'),
        (labels['revenue_growth'], "revenue_growth", BLACK_FONT, None, '0.0%'),
        (f"{labels['cogs']} ({ccy_sym}{unit_suffix})", "cogs", BLUE_FONT, INPUT_FILL, '#,##0'),
        (f"{labels['gross_profit']} ({ccy_sym}{unit_suffix})", "gross_profit", BLACK_FONT, None, '#,##0'),
        (labels['gross_margin'], "gross_margin", BLACK_FONT, None, '0.0%'),
        (f"{labels['rd']} ({ccy_sym}{unit_suffix})", "rd", BLUE_FONT, INPUT_FILL, '#,##0'),
        (labels['rd_ratio'], "rd_ratio", BLACK_FONT, None, '0.0%'),
        (f"{labels['sga']} ({ccy_sym}{unit_suffix})", "sga", BLUE_FONT, INPUT_FILL, '#,##0'),
        (labels['sga_ratio'], "sga_ratio", BLACK_FONT, None, '0.0%'),
        (f"{labels['ebit']} ({ccy_sym}{unit_suffix})", "ebit", BLACK_FONT, None, '#,##0'),
        (labels['ebit_margin'], "ebit_margin", BLACK_FONT, None, '0.0%'),
        (f"{labels['net_income']} ({ccy_sym}{unit_suffix})", "net_income", BLUE_FONT, INPUT_FILL, '#,##0'),
        (labels['net_margin'], "net_margin", BLACK_FONT, None, '0.0%'),
    ]

    for r_idx, (label, key, font, fill, fmt) in enumerate(rows_def, start=2):
        ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER
        for c_idx, y in enumerate(years, start=2):
            val = data[y].get(key)
            if val is not None:
                ws.cell(row=r_idx, column=c_idx, value=val)
            style_data_cell(ws, r_idx, c_idx, font=font, fill=fill, num_fmt=fmt)

    ws.column_dimensions['A'].width = 26
    for c in range(2, len(years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14


def build_valuation_history(ws, val_data, labels, ccy_sym='$'):
    ws.title = labels['sheet_val']
    ws.cell(row=1, column=1, value=labels['label_metric'])
    style_header_row(ws, 1, 1, 1)

    rows = [
        (f"{labels['current_mkt_cap']} ({ccy_sym}B)",
         val_data.get('current_market_cap', 0) / 1e9 if val_data.get('current_market_cap') else None,
         BLUE_FONT, INPUT_FILL, '#,##0.0'),
        (f"{labels['current_price']} ({ccy_sym})",
         val_data.get('current_price'),
         BLUE_FONT, INPUT_FILL, '#,##0.00'),
        (labels['pe_ttm'], val_data.get('current_pe'), BLUE_FONT, INPUT_FILL, '0.0'),
        (labels['pe_fwd'], val_data.get('forward_pe'), BLUE_FONT, INPUT_FILL, '0.0'),
        (f"{labels['shares']} (M)",
         val_data.get('shares_outstanding', 0) / 1e6 if val_data.get('shares_outstanding') else None,
         BLUE_FONT, INPUT_FILL, '#,##0'),
        (f"{labels['wk52_high']} ({ccy_sym})",
         val_data.get('fifty_two_week_high'), BLUE_FONT, INPUT_FILL, '#,##0.00'),
        (f"{labels['wk52_low']} ({ccy_sym})",
         val_data.get('fifty_two_week_low'), BLUE_FONT, INPUT_FILL, '#,##0.00'),
    ]

    for r_idx, (label, val, font, fill, fmt) in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER
        if val is not None:
            ws.cell(row=r_idx, column=2, value=val)
        style_data_cell(ws, r_idx, 2, font=font, fill=fill, num_fmt=fmt)

    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 14


def build_assumptions(ws, assumptions, labels, ccy_sym='$', unit_suffix='B', sources=None):
    """
    构建假设工作表，含三情景并列 + 下拉切换 + 来源注释列（G列）

    sources dict 的每个 key 对应 assumption_rows 中的 label_key，
    值为一段文字说明该假设是怎么推导出来的。
    """
    ws.title = labels['sheet_assumptions']

    ws.cell(row=1, column=1, value=labels['active_scenario'])
    ws.cell(row=1, column=1).font = WHITE_BOLD
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=1).border = THIN_BORDER

    ws.cell(row=1, column=2, value=2)
    ws.cell(row=1, column=2).font = BLACK_BOLD
    ws.cell(row=1, column=2).fill = TOGGLE_FILL
    ws.cell(row=1, column=2).border = THIN_BORDER

    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    dv.error = "1=Bear, 2=Base, 3=Bull"
    dv.errorTitle = "Invalid"
    dv.prompt = "1=Bear, 2=Base, 3=Bull"
    dv.promptTitle = "Select Scenario"
    ws.add_data_validation(dv)
    dv.add(ws["B1"])

    # (label, key, fmt, label_key_for_sources)
    assumption_rows = [
        (f"--- {labels['section_revenue']} ---", None, None, None),
        (f"{labels['tam_start']} ({ccy_sym}{unit_suffix})", "tam_start", '#,##0.0', 'tam_start'),
        (labels['tam_cagr_13'], "tam_cagr_1_3", '0.0%', 'tam_cagr_1_3'),
        (labels['tam_cagr_46'], "tam_cagr_4_6", '0.0%', 'tam_cagr_4_6'),
        (labels['tam_cagr_710'], "tam_cagr_7_10", '0.0%', 'tam_cagr_7_10'),
        (labels['reachable'], "reachable_rate", '0.0%', 'reachable_rate'),
        (labels['share_start'], "share_start", '0.0%', 'share_start'),
        (labels['share_end'], "share_end", '0.0%', 'share_end'),
        (f"--- {labels['section_margin']} ---", None, None, None),
        (labels['gm_start'], "gm_start", '0.0%', 'gm_start'),
        (labels['gm_end'], "gm_end", '0.0%', 'gm_end'),
        (labels['exp_start'], "expense_start", '0.0%', 'expense_start'),
        (labels['exp_end'], "expense_end", '0.0%', 'expense_end'),
        (labels['tax_rate'], "tax_rate", '0.0%', 'tax_rate'),
        (f"--- {labels['section_valuation']} ---", None, None, None),
        (labels['target_pe'], "target_pe", '0.0', 'target_pe'),
        (f"{labels['shares_m']} (M)", "shares_m", '#,##0', 'shares_m'),
    ]

    # 列标题: A=假设项 B=下拉 C=Bear D=Base E=Bull F=活跃值 G=来源
    ws.cell(row=2, column=1, value=labels['label_metric'])
    ws.cell(row=2, column=3, value="Bear")
    ws.cell(row=2, column=4, value="Base")
    ws.cell(row=2, column=5, value="Bull")
    ws.cell(row=2, column=6, value=labels['active_value'])
    ws.cell(row=2, column=7, value=labels['source_label'])
    style_header_row(ws, 2, 1, 7)
    style_header_row(ws, 2, 3, 5, fill=COL_HEADER_FILL, font=BLACK_BOLD)
    style_header_row(ws, 2, 6, 6, fill=OUTPUT_FILL, font=BLACK_BOLD)
    style_header_row(ws, 2, 7, 7,
                     fill=PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid"),
                     font=Font(color="666666", bold=True))

    for r_idx, (label, key, fmt, label_key) in enumerate(assumption_rows, start=3):
        ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER

        if key is None:
            continue

        for c_idx, scenario in enumerate(['bear', 'base', 'bull'], start=3):
            val = assumptions.get(scenario, {}).get(key)
            if val is not None:
                ws.cell(row=r_idx, column=c_idx, value=val)
            style_data_cell(ws, r_idx, c_idx, font=BLUE_FONT, fill=INPUT_FILL, num_fmt=fmt)

        ws.cell(row=r_idx, column=6, value=f'=INDEX(C{r_idx}:E{r_idx},$B$1)')
        style_data_cell(ws, r_idx, 6, font=GREEN_FONT, fill=OUTPUT_FILL, num_fmt=fmt)

        # G列: 来源/推导依据
        source_text = ''
        if sources and label_key in sources:
            source_text = sources[label_key]
        ws.cell(row=r_idx, column=7, value=source_text)
        ws.cell(row=r_idx, column=7).font = GRAY_FONT
        ws.cell(row=r_idx, column=7).fill = NOTE_FILL
        ws.cell(row=r_idx, column=7).border = THIN_BORDER
        ws.cell(row=r_idx, column=7).alignment = WRAP_ALIGN

    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 8
    for c in ['C', 'D', 'E', 'F']:
        ws.column_dimensions[c].width = 14
    ws.column_dimensions['G'].width = 55


def build_projections(ws, assumptions, hist_data, labels, ccy_sym='$',
                       unit_suffix='B', projection_years=10, current_price=None):
    ws.title = labels['sheet_forecast']

    VAL_CURRENT_PRICE = current_price or 100

    hist_years = sorted(hist_data.keys())
    last_hist_year = int(hist_years[-1].replace('A', '')) if hist_years else 2025
    proj_years = [f"{last_hist_year + i + 1}E" for i in range(projection_years)]

    ws.cell(row=1, column=1, value=labels['label_metric'])
    for i, y in enumerate(proj_years):
        ws.cell(row=1, column=i + 2, value=y)
    style_header_row(ws, 1, 1, len(proj_years) + 1)
    style_header_row(ws, 1, 2, len(proj_years) + 1, fill=COL_HEADER_FILL, font=BLACK_BOLD)

    ASSUMPTION_ROWS = {
        'tam_start': 4, 'tam_cagr_1_3': 5, 'tam_cagr_4_6': 6, 'tam_cagr_7_10': 7,
        'reachable_rate': 8, 'share_start': 9, 'share_end': 10,
        'gm_start': 12, 'gm_end': 13, 'expense_start': 14, 'expense_end': 15,
        'tax_rate': 16, 'target_pe': 18, 'shares_m': 19,
    }

    def asum(key):
        row = ASSUMPTION_ROWS[key]
        sn = labels['sheet_assumptions']
        return f"'{sn}'!$F${row}"

    # --- 收入推导 ---
    r = 2
    ws.cell(row=r, column=1, value=f"--- {labels['section_rev_forecast']} ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 3
    ws.cell(row=r, column=1, value=f"AIDC CapEx ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        if i <= 2:
            cagr_ref = asum('tam_cagr_1_3')
        elif i <= 5:
            cagr_ref = asum('tam_cagr_4_6')
        else:
            cagr_ref = asum('tam_cagr_7_10')
        if i == 0:
            ws.cell(row=r, column=col, value=f"={asum('tam_start')}*(1+{cagr_ref})")
        else:
            prev_col = get_column_letter(col - 1)
            ws.cell(row=r, column=col, value=f"={prev_col}{r}*(1+{cagr_ref})")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 4
    ws.cell(row=r, column=1, value=f"{labels['sam']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}3*{asum('reachable_rate')}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 5
    ws.cell(row=r, column=1, value=labels['market_share'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        ws.cell(row=r, column=col,
                value=f"={asum('share_start')}+({asum('share_end')}-{asum('share_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    r = 6
    ws.cell(row=r, column=1, value=f"{labels['fc_revenue']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}4*{c}5")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    r = 7
    ws.cell(row=r, column=1, value=labels['fc_rev_growth'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        if i == 0:
            last_year_data = hist_data.get(hist_years[-1], {}) if hist_years else {}
            hist_rev = last_year_data.get('revenue')
            if hist_rev:
                ws.cell(row=r, column=col,
                        value=f"=({c}6*1000-{hist_rev})/{hist_rev}")
        else:
            prev_col = get_column_letter(col - 1)
            ws.cell(row=r, column=col, value=f"=({c}6-{prev_col}6)/{prev_col}6")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    # --- 利润推导 ---
    r = 9
    ws.cell(row=r, column=1, value=f"--- {labels['section_profit_forecast']} ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 10
    ws.cell(row=r, column=1, value=labels['fc_gm'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2,
                value=f"={asum('gm_start')}+({asum('gm_end')}-{asum('gm_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, i + 2, font=BLACK_FONT, num_fmt='0.0%')

    r = 11
    ws.cell(row=r, column=1, value=f"{labels['fc_gross_profit']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}6*{c}10")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 12
    ws.cell(row=r, column=1, value=labels['fc_expense_ratio'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2,
                value=f"={asum('expense_start')}+({asum('expense_end')}-{asum('expense_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, i + 2, font=BLACK_FONT, num_fmt='0.0%')

    r = 13
    ws.cell(row=r, column=1, value=f"{labels['fc_opex']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}6*{c}12")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 14
    ws.cell(row=r, column=1, value=f"{labels['fc_ebit']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}11-{c}13")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 15
    ws.cell(row=r, column=1, value=labels['fc_tax_rate'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('tax_rate')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='0.0%')

    r = 16
    ws.cell(row=r, column=1, value=f"{labels['fc_net_income']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}14*(1-{c}15)")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    # --- 估值 ---
    r = 18
    ws.cell(row=r, column=1, value=f"--- {labels['section_val_forecast']} ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 19
    ws.cell(row=r, column=1, value=labels['fc_target_pe'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('target_pe')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='0.0')

    r = 20
    ws.cell(row=r, column=1, value=f"{labels['fc_implied_mkt_cap']} ({ccy_sym}{unit_suffix})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}16*{c}19")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    r = 21
    ws.cell(row=r, column=1, value=labels['fc_shares'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('shares_m')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='#,##0')

    r = 22
    ws.cell(row=r, column=1, value=f"{labels['implied_price_label']} ({ccy_sym})")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}20*1000/{c}21")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.00')

    r = 23
    ws.cell(row=r, column=1, value=labels['fc_premium'])
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"=({c}22-{VAL_CURRENT_PRICE})/{VAL_CURRENT_PRICE}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    ws.column_dimensions['A'].width = 26
    for c in range(2, len(proj_years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14


def build_fy_info_sheet(ws, fy_analysis, labels):
    """财年信息工作表：详细说明 FY→自然年映射逻辑"""
    ws.title = labels['sheet_fy_info']

    r = 1
    ws.cell(row=r, column=1, value="Fiscal Year → Calendar Year Mapping")
    ws.cell(row=r, column=1).font = Font(size=14, bold=True, color="1F4E79")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)

    r = 3
    ws.cell(row=r, column=1, value="FY End Month:")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=2, value=fy_analysis.get('fy_end_month', 'N/A'))

    r = 4
    ws.cell(row=r, column=1, value="FY End Date:")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=2, value=fy_analysis.get('fiscal_year_end', 'N/A'))

    r = 5
    is_cal = fy_analysis.get('is_calendar_year', False)
    ws.cell(row=r, column=1, value="Calendar Year Aligned:")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=2, value="Yes" if is_cal else "No — requires mapping")
    ws.cell(row=r, column=2).font = Font(color="FF0000" if not is_cal else "008000", bold=True)

    r = 7
    ws.cell(row=r, column=1, value="Mapping Rule:")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(row=r, column=1, value=fy_analysis.get('mapping_rule', ''))
    ws.cell(row=r, column=1).alignment = WRAP_ALIGN

    r = 9
    ws.cell(row=r, column=1, value="Detailed Mappings:")
    ws.cell(row=r, column=1).font = Font(size=12, bold=True, color="1F4E79")

    r = 10
    ws.cell(row=r, column=1, value="Report Date")
    ws.cell(row=r, column=2, value="FY Label")
    ws.cell(row=r, column=3, value="Natural Year")
    style_header_row(ws, r, 1, 3)

    mappings = fy_analysis.get('date_mappings', {})
    fy_end_month = fy_analysis.get('fy_end_month', 12)

    for i, (date_str, nat_year) in enumerate(sorted(mappings.items())):
        row = r + 1 + i
        if fy_end_month <= 6:
            fy_label = f"FY{nat_year + 1}"
        else:
            fy_label = f"FY{nat_year}"

        ws.cell(row=row, column=1, value=date_str)
        ws.cell(row=row, column=2, value=fy_label)
        ws.cell(row=row, column=3, value=f"{nat_year}A")
        for c in range(1, 4):
            ws.cell(row=row, column=c).border = THIN_BORDER

    note_row = r + 1 + len(mappings) + 2
    ws.cell(row=note_row, column=1, value="Full Notes:")
    ws.cell(row=note_row, column=1).font = BLACK_BOLD
    note = fy_analysis.get('fiscal_year_note', '')
    ws.cell(row=note_row + 1, column=1, value=note)
    ws.cell(row=note_row + 1, column=1).alignment = WRAP_ALIGN
    ws.merge_cells(start_row=note_row + 1, start_column=1, end_row=note_row + 3, end_column=4)

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16


def build_sanity_checks(ws, checks, labels):
    ws.title = labels['sheet_sanity']

    ws.cell(row=1, column=1, value=labels['label_metric'])
    ws.cell(row=1, column=2, value="Bear")
    ws.cell(row=1, column=3, value="Base")
    ws.cell(row=1, column=4, value="Bull")
    ws.cell(row=1, column=5, value="Status")
    style_header_row(ws, 1, 1, 5)

    for r_idx, check in enumerate(checks, start=2):
        ws.cell(row=r_idx, column=1, value=check['name'])
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER
        for c_idx, scenario in enumerate(['bear', 'base', 'bull'], start=2):
            ws.cell(row=r_idx, column=c_idx, value=check.get(scenario, ''))
            style_data_cell(ws, r_idx, c_idx)
        ws.cell(row=r_idx, column=5, value=check.get('status', ''))
        style_data_cell(ws, r_idx, 5)

    ws.column_dimensions['A'].width = 30
    for c in ['B', 'C', 'D', 'E']:
        ws.column_dimensions[c].width = 14


def main():
    parser = argparse.ArgumentParser(description='生成估值模型 Excel')
    parser.add_argument('--ticker', required=True, help='股票代码')
    parser.add_argument('--input', help='模型数据 JSON 文件路径')
    parser.add_argument('--output-dir', default='./out', help='输出目录')
    parser.add_argument('--projection-years', type=int, default=10, help='预测年数')
    parser.add_argument('--language', choices=['zh-CN', 'en'], default=None,
                        help='输出语言（默认自动检测）')
    args = parser.parse_args()

    data = {}
    val_data = {}
    assumptions = {}
    checks = []
    sources = {}
    fy_analysis = {}
    language = args.language or 'en'
    ccy_sym = '$'
    unit_suffix = 'M'

    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            raw = json.load(f)
            data = raw.get('hist_data', raw.get('data', {}))
            val_data = raw.get('val_data', {})
            assumptions = raw.get('assumptions', {})
            checks = raw.get('checks', [])
            sources = raw.get('sources', {})
            fy_analysis = raw.get('fiscal_year_analysis', {})
            if not args.language:
                language = raw.get('language', 'en')
            ccy_sym = raw.get('currency_symbol', '$')
            unit_suffix = raw.get('unit_suffix', 'M')

    labels = get_labels(language)

    wb = Workbook()

    ws1 = wb.active
    build_financial_history(ws1, data, labels, ccy_sym, unit_suffix)

    ws2 = wb.create_sheet()
    build_valuation_history(ws2, val_data, labels, ccy_sym)

    if fy_analysis:
        ws_fy = wb.create_sheet()
        build_fy_info_sheet(ws_fy, fy_analysis, labels)

    if assumptions:
        ws3 = wb.create_sheet()
        build_assumptions(ws3, assumptions, labels, ccy_sym, 'B', sources)

        ws4 = wb.create_sheet()
        current_price = val_data.get('current_price')
        build_projections(ws4, assumptions, data, labels, ccy_sym, 'B',
                          args.projection_years, current_price=current_price)

        if checks:
            ws5 = wb.create_sheet()
            build_sanity_checks(ws5, checks, labels)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{args.ticker}_coverage_model.xlsx"
    wb.save(str(output_path))

    print(f"模型已生成: {output_path}")
    return str(output_path)


if __name__ == '__main__':
    main()
