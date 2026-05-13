#!/usr/bin/env python3
"""
build_excel.py — 生成带场景切换的估值模型 Excel
用法: python build_excel.py --ticker NVDA --input data.json

输入: fetch_financials.py 输出的 JSON 文件（或直接传入参数）
输出: ./out/<ticker>_coverage_model.xlsx

依赖: openpyxl
"""

import sys
import json
import argparse
import os
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


# ===== 颜色定义 =====
BLUE_FONT = Font(color="0000FF")          # 硬编码输入
BLACK_FONT = Font(color="000000")         # 公式
GREEN_FONT = Font(color="008000")         # 跨表链接
WHITE_BOLD = Font(color="FFFFFF", bold=True)
BLACK_BOLD = Font(color="000000", bold=True)

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")  # 深蓝
COL_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")  # 浅蓝
OUTPUT_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")  # 中蓝
INPUT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # 浅灰
TOGGLE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # 浅黄

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

RIGHT_ALIGN = Alignment(horizontal='right')


def style_header_row(ws, row, start_col, end_col, fill=HEADER_FILL, font=WHITE_BOLD):
    """设置表头行样式"""
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center')
        cell.border = THIN_BORDER


def style_data_cell(ws, row, col, font=BLACK_FONT, fill=None, num_fmt=None):
    """设置数据单元格样式"""
    cell = ws.cell(row=row, column=col)
    cell.font = font
    if fill:
        cell.fill = fill
    cell.border = THIN_BORDER
    cell.alignment = RIGHT_ALIGN
    if num_fmt:
        cell.number_format = num_fmt


def build_financial_history(ws, data: dict):
    """构建财务历史工作表"""
    ws.title = "财务历史"

    years = sorted(data.keys())
    if not years:
        return

    # 表头
    ws.cell(row=1, column=1, value="指标")
    style_header_row(ws, 1, 1, len(years) + 1)
    for i, y in enumerate(years):
        ws.cell(row=1, column=i + 2, value=y)
        style_header_row(ws, 1, i + 2, i + 2, fill=COL_HEADER_FILL, font=BLACK_BOLD)

    # 数据行定义
    rows_def = [
        ("收入 ($M)", "revenue", BLUE_FONT, INPUT_FILL, '#,##0'),
        ("收入增速", "revenue_growth", BLACK_FONT, None, '0.0%'),
        ("营业成本 ($M)", "cogs", BLUE_FONT, INPUT_FILL, '#,##0'),
        ("毛利 ($M)", "gross_profit", BLACK_FONT, None, '#,##0'),
        ("毛利率", "gross_margin", BLACK_FONT, None, '0.0%'),
        ("研发费用 ($M)", "rd", BLUE_FONT, INPUT_FILL, '#,##0'),
        ("研发费用率", "rd_ratio", BLACK_FONT, None, '0.0%'),
        ("销售管理费用 ($M)", "sga", BLUE_FONT, INPUT_FILL, '#,##0'),
        ("销售管理费用率", "sga_ratio", BLACK_FONT, None, '0.0%'),
        ("EBIT ($M)", "ebit", BLACK_FONT, None, '#,##0'),
        ("EBIT margin", "ebit_margin", BLACK_FONT, None, '0.0%'),
        ("净利润 ($M)", "net_income", BLUE_FONT, INPUT_FILL, '#,##0'),
        ("净利率", "net_margin", BLACK_FONT, None, '0.0%'),
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

    # 列宽
    ws.column_dimensions['A'].width = 22
    for c in range(2, len(years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14


def build_valuation_history(ws, val_data: dict):
    """构建估值历史工作表"""
    ws.title = "估值历史"

    ws.cell(row=1, column=1, value="指标")
    style_header_row(ws, 1, 1, 1)

    rows = [
        ("当前市值 ($B)", val_data.get('current_market_cap', 0) / 1e9 if val_data.get('current_market_cap') else None, BLUE_FONT, INPUT_FILL, '#,##0.0'),
        ("当前股价 ($)", val_data.get('current_price'), BLUE_FONT, INPUT_FILL, '#,##0.00'),
        ("PE (TTM)", val_data.get('current_pe'), BLUE_FONT, INPUT_FILL, '0.0'),
        ("PE (Forward)", val_data.get('forward_pe'), BLUE_FONT, INPUT_FILL, '0.0'),
        ("流通股数 (M)", val_data.get('shares_outstanding', 0) / 1e6 if val_data.get('shares_outstanding') else None, BLUE_FONT, INPUT_FILL, '#,##0'),
        ("52周高 ($)", val_data.get('fifty_two_week_high'), BLUE_FONT, INPUT_FILL, '#,##0.00'),
        ("52周低 ($)", val_data.get('fifty_two_week_low'), BLUE_FONT, INPUT_FILL, '#,##0.00'),
    ]

    for r_idx, (label, val, font, fill, fmt) in enumerate(rows, start=2):
        ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER
        if val is not None:
            ws.cell(row=r_idx, column=2, value=val)
        style_data_cell(ws, r_idx, 2, font=font, fill=fill, num_fmt=fmt)

    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 14


def build_assumptions(ws, assumptions: dict):
    """
    构建假设工作表，含三情景并列 + 下拉切换

    assumptions 结构:
    {
        'bear': {'tam_start': ..., 'tam_cagr_1_3': ..., ...},
        'base': {...},
        'bull': {...},
    }
    """
    ws.title = "假设"

    # ===== 场景切换控件 =====
    ws.cell(row=1, column=1, value="活跃情景")
    ws.cell(row=1, column=1).font = WHITE_BOLD
    ws.cell(row=1, column=1).fill = HEADER_FILL
    ws.cell(row=1, column=1).border = THIN_BORDER

    ws.cell(row=1, column=2, value=2)  # 默认Base
    ws.cell(row=1, column=2).font = BLACK_BOLD
    ws.cell(row=1, column=2).fill = TOGGLE_FILL
    ws.cell(row=1, column=2).border = THIN_BORDER

    # 添加下拉验证
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    dv.error = "请选择 1(Bear), 2(Base), 3(Bull)"
    dv.errorTitle = "无效输入"
    dv.prompt = "1=Bear, 2=Base, 3=Bull"
    dv.promptTitle = "选择情景"
    ws.add_data_validation(dv)
    dv.add(ws["B1"])

    # ===== 假设行定义 =====
    assumption_rows = [
        # (标签, key, 单位格式)
        ("--- 收入驱动 ---", None, None),
        ("TAM 起始 ($B)", "tam_start", '#,##0.0'),
        ("TAM CAGR Y1-Y3", "tam_cagr_1_3", '0.0%'),
        ("TAM CAGR Y4-Y6", "tam_cagr_4_6", '0.0%'),
        ("TAM CAGR Y7-Y10", "tam_cagr_7_10", '0.0%'),
        ("可触达率", "reachable_rate", '0.0%'),
        ("起始市占率", "share_start", '0.0%'),
        ("终年市占率", "share_end", '0.0%'),
        ("--- 利润率 ---", None, None),
        ("起始毛利率", "gm_start", '0.0%'),
        ("终年毛利率", "gm_end", '0.0%'),
        ("起始费用率", "expense_start", '0.0%'),
        ("终年费用率", "expense_end", '0.0%'),
        ("税率", "tax_rate", '0.0%'),
        ("--- 估值 ---", None, None),
        ("目标 PE", "target_pe", '0.0'),
        ("当前股份数 (M)", "shares_m", '#,##0'),
    ]

    # 列标题
    ws.cell(row=2, column=1, value="假设项")
    ws.cell(row=2, column=3, value="Bear")
    ws.cell(row=2, column=4, value="Base")
    ws.cell(row=2, column=5, value="Bull")
    ws.cell(row=2, column=6, value="活跃值")
    style_header_row(ws, 2, 1, 6)
    style_header_row(ws, 2, 3, 5, fill=COL_HEADER_FILL, font=BLACK_BOLD)
    style_header_row(ws, 2, 6, 6, fill=OUTPUT_FILL, font=BLACK_BOLD)

    for r_idx, (label, key, fmt) in enumerate(assumption_rows, start=3):
        ws.cell(row=r_idx, column=1, value=label)
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER

        if key is None:
            # 分隔行
            continue

        # 三情景数值
        for c_idx, scenario in enumerate(['bear', 'base', 'bull'], start=3):
            val = assumptions.get(scenario, {}).get(key)
            if val is not None:
                ws.cell(row=r_idx, column=c_idx, value=val)
            style_data_cell(ws, r_idx, c_idx, font=BLUE_FONT, fill=INPUT_FILL, num_fmt=fmt)

        # INDEX 公式：活跃值 = INDEX(Bear:Bull, $B$1)
        ws.cell(row=r_idx, column=6, value=f'=INDEX(C{r_idx}:E{r_idx},$B$1)')
        style_data_cell(ws, r_idx, 6, font=GREEN_FONT, fill=OUTPUT_FILL, num_fmt=fmt)

    ws.column_dimensions['A'].width = 22
    for c in ['B', 'C', 'D', 'E', 'F']:
        ws.column_dimensions[c].width = 14


def build_projections(ws, assumptions: dict, hist_data: dict, projection_years: int = 10,
                      current_price: float = None):
    """
    构建预测工作表

    从假设表拉取参数，生成收入→利润→估值推导
    """
    ws.title = "预测"

    VAL_CURRENT_PRICE = current_price or 100  # 默认值，实际从外部传入

    # 确定起始年份
    hist_years = sorted(hist_data.keys())
    last_hist_year = int(hist_years[-1].replace('A', '')) if hist_years else 2025
    proj_years = [f"{last_hist_year + i + 1}E" for i in range(projection_years)]

    # 表头
    ws.cell(row=1, column=1, value="指标")
    for i, y in enumerate(proj_years):
        ws.cell(row=1, column=i + 2, value=y)
    style_header_row(ws, 1, 1, len(proj_years) + 1)
    style_header_row(ws, 1, 2, len(proj_years) + 1, fill=COL_HEADER_FILL, font=BLACK_BOLD)

    # 行定义：row -> (标签, 引用假设key的公式模板)
    # 注意：这里使用字符串引用 '假设' sheet 的单元格
    # 实际公式需要知道假设表中各参数的行号

    # 为了简化，直接用假设 sheet 的列F（活跃值）的引用
    # 假设表中：F3=TAM起始, F4=TAM CAGR Y1-3, F5=TAM CAGR Y4-6, ...
    # 行号对应 build_assumptions 中的 assumption_rows

    # 对应假设表的行号 (从第3行开始)
    ASSUMPTION_ROWS = {
        'tam_start': 4,      # 第4行
        'tam_cagr_1_3': 5,
        'tam_cagr_4_6': 6,
        'tam_cagr_7_10': 7,
        'reachable_rate': 8,
        'share_start': 9,
        'share_end': 10,
        'gm_start': 12,
        'gm_end': 13,
        'expense_start': 14,
        'expense_end': 15,
        'tax_rate': 16,
        'target_pe': 18,
        'shares_m': 19,
    }

    def asum(key):
        """生成对假设表活跃值列的引用"""
        row = ASSUMPTION_ROWS[key]
        return f"假设!$F${row}"

    # === 收入推导 ===
    r = 2
    section_headers = [
        (r, "收入推导"),  # 2
    ]
    ws.cell(row=r, column=1, value="--- 收入推导 ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 3  # TAM
    ws.cell(row=r, column=1, value="TAM ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c_letter = get_column_letter(col)
        # 分段 CAGR — 第一年也应用增长
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

    r = 4  # 可触达市场
    ws.cell(row=r, column=1, value="可触达市场 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c_letter = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c_letter}3*{asum('reachable_rate')}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 5  # 市占率 (线性插值)
    ws.cell(row=r, column=1, value="市占率")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        # 线性插值：start + (end - start) * (i / (n-1))
        ws.cell(row=r, column=col,
                value=f"={asum('share_start')}+({asum('share_end')}-{asum('share_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    r = 6  # 收入
    ws.cell(row=r, column=1, value="收入 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c_letter = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c_letter}4*{c_letter}5")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    r = 7  # 收入增速
    ws.cell(row=r, column=1, value="收入增速")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c_letter = get_column_letter(col)
        if i == 0:
            # 第一年增速 vs 历史最后一年（需要知道历史收入）
            hist_rev = None
            last_year_data = hist_data.get(hist_years[-1], {}) if hist_years else {}
            hist_rev = last_year_data.get('revenue')
            if hist_rev:
                # 历史收入单位是$M，预测是$B，需要转换
                ws.cell(row=r, column=col,
                        value=f"=({c_letter}6*1000-{hist_rev})/{hist_rev}")
        else:
            prev_col = get_column_letter(col - 1)
            ws.cell(row=r, column=col, value=f"=({c_letter}6-{prev_col}6)/{prev_col}6")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    # === 利润推导 ===
    r = 9
    ws.cell(row=r, column=1, value="--- 利润推导 ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 10  # 毛利率 (线性插值)
    ws.cell(row=r, column=1, value="毛利率")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2,
                value=f"={asum('gm_start')}+({asum('gm_end')}-{asum('gm_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, i + 2, font=BLACK_FONT, num_fmt='0.0%')

    r = 11  # 毛利
    ws.cell(row=r, column=1, value="毛利 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}6*{c}10")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 12  # 费用率 (线性插值)
    ws.cell(row=r, column=1, value="费用率")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2,
                value=f"={asum('expense_start')}+({asum('expense_end')}-{asum('expense_start')})*{i}/{projection_years - 1}")
        style_data_cell(ws, r, i + 2, font=BLACK_FONT, num_fmt='0.0%')

    r = 13  # 运营费用
    ws.cell(row=r, column=1, value="运营费用 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}6*{c}12")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 14  # EBIT
    ws.cell(row=r, column=1, value="EBIT ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}11-{c}13")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='#,##0.0')

    r = 15  # 税率
    ws.cell(row=r, column=1, value="税率")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('tax_rate')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='0.0%')

    r = 16  # 净利润
    ws.cell(row=r, column=1, value="净利润 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}14*(1-{c}15)")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    # === 估值 ===
    r = 18
    ws.cell(row=r, column=1, value="--- 估值 ---")
    ws.cell(row=r, column=1).font = WHITE_BOLD
    ws.cell(row=r, column=1).fill = HEADER_FILL
    for c in range(2, len(proj_years) + 2):
        ws.cell(row=r, column=c).fill = HEADER_FILL

    r = 19  # 目标PE (固定值)
    ws.cell(row=r, column=1, value="目标 PE")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('target_pe')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='0.0')

    r = 20  # 隐含市值
    ws.cell(row=r, column=1, value="隐含市值 ($B)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"={c}16*{c}19")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.0')

    r = 21  # 股份数
    ws.cell(row=r, column=1, value="股份数 (M)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        ws.cell(row=r, column=i + 2, value=f"={asum('shares_m')}")
        style_data_cell(ws, r, i + 2, font=GREEN_FONT, num_fmt='#,##0')

    r = 22  # 隐含股价
    ws.cell(row=r, column=1, value="隐含股价 ($)")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        # 市值$B / 股份数M = 市值$B * 1000 / 股份数M = $
        ws.cell(row=r, column=col, value=f"={c}20*1000/{c}21")
        style_data_cell(ws, r, col, font=BLACK_FONT, fill=OUTPUT_FILL, num_fmt='#,##0.00')

    r = 23  # vs 当前股价溢价
    ws.cell(row=r, column=1, value="vs 当前股价溢价")
    ws.cell(row=r, column=1).font = BLACK_BOLD
    ws.cell(row=r, column=1).border = THIN_BORDER
    for i in range(projection_years):
        col = i + 2
        c = get_column_letter(col)
        ws.cell(row=r, column=col, value=f"=({c}22-{VAL_CURRENT_PRICE})/{VAL_CURRENT_PRICE}")
        style_data_cell(ws, r, col, font=BLACK_FONT, num_fmt='0.0%')

    # 列宽
    ws.column_dimensions['A'].width = 22
    for c in range(2, len(proj_years) + 2):
        ws.column_dimensions[get_column_letter(c)].width = 14


def build_sanity_checks(ws, checks: list):
    """构建合理性检验工作表"""
    ws.title = "合理性"

    ws.cell(row=1, column=1, value="检验项")
    ws.cell(row=1, column=2, value="Bear")
    ws.cell(row=1, column=3, value="Base")
    ws.cell(row=1, column=4, value="Bull")
    ws.cell(row=1, column=5, value="状态")
    style_header_row(ws, 1, 1, 5)

    for r_idx, check in enumerate(checks, start=2):
        ws.cell(row=r_idx, column=1, value=check['name'])
        ws.cell(row=r_idx, column=1).font = BLACK_BOLD
        ws.cell(row=r_idx, column=1).border = THIN_BORDER

        for c_idx, scenario in enumerate(['bear', 'base', 'bull'], start=2):
            val = check.get(scenario, '')
            ws.cell(row=r_idx, column=c_idx, value=val)
            style_data_cell(ws, r_idx, c_idx)

        # 状态
        status = check.get('status', '')
        ws.cell(row=r_idx, column=5, value=status)
        style_data_cell(ws, r_idx, 5)

    ws.column_dimensions['A'].width = 30
    for c in ['B', 'C', 'D', 'E']:
        ws.column_dimensions[c].width = 14


def main():
    parser = argparse.ArgumentParser(description='生成估值模型 Excel')
    parser.add_argument('--ticker', required=True, help='股票代码')
    parser.add_argument('--input', help='模型数据 JSON 文件路径（googl_model_data.py 输出）')
    parser.add_argument('--output-dir', default='./out', help='输出目录')
    parser.add_argument('--projection-years', type=int, default=10, help='预测年数')
    args = parser.parse_args()

    # 加载数据
    data = {}
    val_data = {}
    assumptions = {}
    checks = []

    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            raw = json.load(f)
            data = raw.get('hist_data', raw.get('data', {}))
            val_data = raw.get('val_data', {})
            assumptions = raw.get('assumptions', {})
            checks = raw.get('checks', [])

    # 创建工作簿
    wb = Workbook()

    # Sheet 1: 财务历史
    ws1 = wb.active
    build_financial_history(ws1, data)

    # Sheet 2: 估值历史
    ws2 = wb.create_sheet()
    build_valuation_history(ws2, val_data)

    # Sheet 3: 假设
    if assumptions:
        ws3 = wb.create_sheet()
        build_assumptions(ws3, assumptions)

        # Sheet 4: 预测
        ws4 = wb.create_sheet()
        current_price = val_data.get('current_price')
        build_projections(ws4, assumptions, data, args.projection_years,
                          current_price=current_price)

        # Sheet 5: 合理性
        if checks:
            ws5 = wb.create_sheet()
            build_sanity_checks(ws5, checks)

    # 保存
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{args.ticker}_coverage_model.xlsx"
    wb.save(str(output_path))

    print(f"模型已生成: {output_path}")
    return str(output_path)


if __name__ == '__main__':
    main()
