#!/usr/bin/env python3
"""构建 MiniMax (00100.HK) 估值模型数据 JSON"""
import json
import os

os.makedirs('out', exist_ok=True)

data = {
    "ticker": "00100.HK",
    "company_name": "MiniMax Group (稀宇科技)",
    "language": "zh-CN",
    "currency_symbol": "HK$",
    "unit_suffix": "M",
    "notes": "MiniMax以美元计收入(海外收入>70%),但Excel以港元为显示货币。模型假设基于美元,市值/股价以港元计。1 USD ≈ 7.8 HKD",

    "hist_data": {
        "2023A": {
            "revenue": 3.46,
            "revenue_growth": None,
            "cogs": 4.315,
            "gross_profit": -0.855,
            "gross_margin": -0.247,
            "rd": 5.0,
            "rd_ratio": 1.445,
            "sga": 2.0,
            "sga_ratio": 0.578,
            "ebit": -7.855,
            "ebit_margin": -2.270,
            "net_income": -8.0,
            "net_margin": -2.312
        },
        "2024A": {
            "revenue": 30.523,
            "revenue_growth": 7.820,
            "cogs": 26.785,
            "gross_profit": 3.738,
            "gross_margin": 0.122,
            "rd": 188.979,
            "rd_ratio": 6.191,
            "sga": 86.995,
            "sga_ratio": 2.850,
            "ebit": -286.620,
            "ebit_margin": -9.389,
            "net_income": -465.238,
            "net_margin": -15.242
        },
        "2025A": {
            "revenue": 79.038,
            "revenue_growth": 1.589,
            "cogs": 58.959,
            "gross_profit": 20.079,
            "gross_margin": 0.254,
            "rd": 252.771,
            "rd_ratio": 3.198,
            "sga": 88.709,
            "sga_ratio": 1.122,
            "ebit": -281.032,
            "ebit_margin": -3.555,
            "net_income": -1871.617,
            "net_margin": -23.681
        }
    },

    "segment_history": {},

    "val_data": {
        "current_market_cap": 232000000000,
        "current_price": 735.0,
        "current_pe": -14.9,
        "forward_pe": -19.0,
        "shares_outstanding": 309300000,
        "fifty_two_week_high": 1330.0,
        "fifty_two_week_low": 220.0
    },

    "fiscal_year_analysis": {
        "fy_end_month": 12,
        "fiscal_year_end": "12/31",
        "is_calendar_year": True,
        "mapping_rule": "FY = 自然年。MiniMax财年结束于12月31日,与自然年一致。2023A/2024A/2025A均对应自然年度。",
        "fiscal_year_note": "MiniMax于2026年1月9日在港交所上市,2025年年报为首份上市后年报。2023-2024年为上市前财务数据(来自招股书)。所有数据已根据IFRS编制。2023年收入$3.46M,2024年$30.52M(+782%),2025年$79.04M(+159%)。2025年净亏损$1,872M含$1,590M公允价值变动(非现金),调整后净亏损$251M。"
    },

    "assumptions": {
        "bear": {
            "tam_start": 48.0,
            "tam_cagr_1_3": 0.20,
            "tam_cagr_4_6": 0.12,
            "tam_cagr_7_10": 0.08,
            "reachable_rate": 0.50,
            "share_start": 0.0012,
            "share_end": 0.008,
            "gm_start": 0.23,
            "gm_end": 0.40,
            "expense_start": 3.50,
            "expense_end": 0.80,
            "tax_rate": 0.15,
            "pe_start": 15.0,
            "pe_end": 20.0,
            "shares_m": 309.3
        },
        "base": {
            "tam_start": 48.0,
            "tam_cagr_1_3": 0.30,
            "tam_cagr_4_6": 0.20,
            "tam_cagr_7_10": 0.15,
            "reachable_rate": 0.60,
            "share_start": 0.0016,
            "share_end": 0.025,
            "gm_start": 0.25,
            "gm_end": 0.60,
            "expense_start": 3.00,
            "expense_end": 0.50,
            "tax_rate": 0.15,
            "pe_start": 20.0,
            "pe_end": 30.0,
            "shares_m": 309.3
        },
        "bull": {
            "tam_start": 48.0,
            "tam_cagr_1_3": 0.40,
            "tam_cagr_4_6": 0.28,
            "tam_cagr_7_10": 0.20,
            "reachable_rate": 0.70,
            "share_start": 0.0020,
            "share_end": 0.050,
            "gm_start": 0.27,
            "gm_end": 0.75,
            "expense_start": 2.80,
            "expense_end": 0.35,
            "tax_rate": 0.15,
            "pe_start": 25.0,
            "pe_end": 40.0,
            "shares_m": 309.3
        }
    },

    "sources": {
        "tam_start": (
            "【数据】IDC报告:2025年全球AI大模型市场规模约480亿美元(2026年2月)。"
            "Grand View Research:2024年中国AI API市场35.6亿美元。"
            "JPMorgan预测2030年全球AI市场1.4万亿美元。"
            "【逻辑】取IDC的全球AI大模型市场480亿美元为TAM起点,涵盖API调用、模型订阅、AI原生产品。"
            "中国大模型核心产业规模>1.2万亿元(约$165B),含硬件/算力,去除后API+订阅约$48B合理。"
            "【情景差异】Bear/Base/Bull共用同一TAM起点,差异体现在CAGR和市占率。"
        ),
        "tam_cagr_1_3": (
            "【数据】Grand View Research:中国AI API市场CAGR 32.4%(2024-2030)。"
            "IDC:生成式AI计算市场从2022年8.2亿美元增至2026年109.9亿美元。"
            "【逻辑】Base Y1-Y3 CAGR 30%反映市场高速增长期,AI应用加速渗透。"
            "Bear 20%假设技术成熟度低于预期、企业IT支出放缓。Bull 40%假设Agent/AI原生应用爆发式增长。"
        ),
        "tam_cagr_4_6": (
            "【数据】Gartner预测AI软件市场2025-2028 CAGR约22%。"
            "【逻辑】Base Y4-Y6 CAGR 20%反映市场从超高速转入高速增长。增速自然衰减。"
            "Bear 12%假设AI投资回报率受质疑、监管收紧。Bull 28%假设AI全面渗透实体经济。"
        ),
        "tam_cagr_7_10": (
            "【数据】PwC预测2030年AI贡献全球GDP 15.7万亿美元。"
            "【逻辑】Base Y7-Y10 CAGR 15%,市场逐渐成熟但仍有增长空间。"
            "Bear 8%接近GDP增速,假设AI技术不再有实质性突破。Bull 20%维持高增速。"
        ),
        "reachable_rate": (
            "【数据】MiniMax海外收入>70%,覆盖200+国家/地区。Token消耗占全球4.2%(OpenRouter 2026年2月数据)。"
            "【逻辑】Base 60%考虑:中国+一带一路+部分欧美市场可触达。美国市场受地缘政治限制,可能被排除。"
            "Bear 50%假设更严格的技术出口管制。Bull 70%假设MiniMax全球化成功突破壁垒。"
        ),
        "share_start": (
            "【数据】MiniMax 2025年收入$79M vs 全球AI模型市场$48B,市占率约0.165%。"
            "OpenRouter Token消耗份额4.2%说明用量份额高于收入份额(定价策略低价抢量)。"
            "【逻辑】以收入市占率为准。Base 0.16%接近实际水平。Bear 0.12%保守。Bull 0.20%乐观。"
        ),
        "share_end": (
            "【数据】高盛(Goldman Sachs)2026年3月报告预测MiniMax 2030年全球市占率2.5%。"
            "OpenAI 2025年占全球AI市场约25%+($131B收入/$480B TAM),Anthropic约8-10%。"
            "【逻辑】Base 2.5%对标高盛预测,MiniMax以API平台定位逐步提升份额。"
            "Bear 0.8%假设大厂挤压独立AI公司。Bull 5.0%假设MiniMax成为全球前5大AI平台。"
        ),
        "gm_start": (
            "【数据】MiniMax 2025年毛利率25.4%(官方年报2026年3月2日发布)。"
            "B2B API业务毛利率69.4%,B2C(Talkie/海螺AI)拉低整体毛利率。"
            "【逻辑】Base起始25%,接近实际值。随着API占比提升,毛利率改善。"
            "Bear 23%假设B2C烧钱更久。Bull 27%假设API收入占比快速提升。"
        ),
        "gm_end": (
            "【数据】高盛预测MiniMax API毛利率可持续提升。Anthropic毛利率约55-65%。"
            "成熟SaaS/云平台毛利率通常60-80%。Nvidia软件毛利率~75%。"
            "【逻辑】Base终年60%参考成熟AI平台公司水平。"
            "Bear 40%假设竞争压价。Bull 75%对标顶级AI基础设施公司。"
        ),
        "expense_start": (
            "【数据】MiniMax 2025年运营费用:R&D $252.8M+S&D $51.9M+Admin $36.8M = $341.5M。"
            "收入$79M,费用率=341.5/79=432%。净其他收入$40.4M部分抵消。"
            "【逻辑】Base起始300%反映2026年费用率开始改善(收入高增长摊薄固定成本)。"
            "Bear 350%假设费用改善更慢。Bull 280%假设收入爆发式增长拉动。"
        ),
        "expense_end": (
            "【数据】高盛预测MiniMax 2035年调整后EBIT利润率21%。"
            "成熟软件公司费用率通常40-60%(R&D 15-25%+S&M 15-25%+G&A 5-10%)。"
            "【逻辑】Base终年50%:R&D 20%+S&M 20%+Admin 10%。"
            "Bear 80%假设长期高研发投入。Bull 35%假设经营杠杆极致释放。"
        ),
        "tax_rate": (
            "【数据】MiniMax注册于开曼群岛,运营主体在中国(25%法定税率)。"
            "但作为高新技术企业可享15%优惠税率。港股上市主体在境外,有效税率较低。"
            "【逻辑】Base 15%统一假设。MiniMax目前没有所得税费用(亏损状态)。"
        ),
        "pe_start": (
            "【数据】MiniMax当前PE为负(TTM -14.9x),不适用。"
            "可比公司:OpenAI PS ~34x(未盈利),Anthropic PS ~13-27x(未盈利)。"
            "传统科技公司:Microsoft PS ~12x PE ~35x,Google PS ~6x PE ~22x。"
            "【逻辑】起始PE给20x(Bear 15/Bull 25),反映AI赛道高成长溢价但尚未盈利。"
        ),
        "pe_end": (
            "【数据】高增长科技公司成熟后PE通常20-40x。Salesforce ~28x, ServiceNow ~55x。"
            "【逻辑】Base终年PE 30x:若MiniMax 2035年盈利稳定增长,享有AI平台溢价。"
            "Bear 20x等同于普通软件公司。Bull 40x类似顶级SaaS平台。"
        ),
        "shares_m": (
            "【数据】Investing.com:MiniMax流通股309.3M(截至2026年5月)。"
            "雪球显示总股本3.14亿,港股股本2.33亿。"
            "【逻辑】取309.3M为基准。不考虑未来增发/回购(不确定性高)。"
            "上市时14名基石投资者锁定67.88%股份,锁定期6-12个月。"
        )
    },

    "checks": []
}

output_path = 'out/minimax_model_data.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f'Model data saved to {output_path}')
print(f'File size: {os.path.getsize(output_path):,} bytes')

# Verify
with open(output_path, 'r', encoding='utf-8') as f:
    verify = json.load(f)
print(f'Verification: {len(verify.keys())} top-level keys')
print(f'Hist years: {sorted(verify["hist_data"].keys())}')
print(f'Scenarios: {sorted(verify["assumptions"].keys())}')
print(f'Current price: {verify["val_data"]["current_price"]}')
print(f'Current market cap: {verify["val_data"]["current_market_cap"]:,}')
