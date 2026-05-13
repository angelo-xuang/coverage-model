# Coverage Model — 个人投资估值模型

构建个人可理解、可控制的股权估值模型。核心哲学：**人做4个定性判断，AI做180+个定量映射**。

## 工作流程

```
Phase 1: 数据拉取 → Phase 2: 背景呈现 → Phase 3: 讨论定性判断 → Phase 4: 建模执行 → Phase 5: 合理性检验 → Phase 6: Excel输出
```

## 项目结构

```
coverage-model/
├── SKILL.md              # 完整工作流定义（Claude Skill）
├── references/
│   ├── methodology.md    # 定性→定量映射表、TAM预测方法
│   ├── formatting.md     # Excel格式规范
│   └── sanity-checks.md  # 三层合理性检验框架
├── scripts/
│   ├── fetch_financials.py   # 拉取5年损益表数据
│   ├── fetch_segments.py     # 拉取业务分部数据
│   ├── build_excel.py        # 生成估值模型Excel
│   └── googl_model_data.py   # Google模型数据示例
└── out/                      # 输出目录
```

## 快速开始

```bash
# 拉取财务数据
python scripts/fetch_financials.py NVDA --proxy 127.0.0.1:7890

# 拉取分部数据
python scripts/fetch_segments.py NVDA --proxy 127.0.0.1:7890

# 生成Excel模型
python scripts/build_excel.py --ticker GOOGL --input out/googl_model_data.json
```

## 特性

- **自上而下 TAM 方法**：从行业总量出发推导公司收入
- **三情景对比**：Bear / Base / Bull 自动推导
- **Excel 场景切换**：下拉菜单切换活跃情景，INDEX公式联动
- **合理性检验**：三层验证（内部逻辑、跨情景排序、外部基准）
- **财年自动映射**：自动将 FY 转换为自然年
- **多币种支持**：根据公司运营主体自动选择货币单位和语言

## 依赖

```
yfinance
openpyxl
requests
```

## 数据源

| 数据 | 来源 |
|------|------|
| 财务报表 | yfinance + SEC EDGAR |
| 分部数据 | SEC EDGAR 10-K |
| 市场事件 | WebSearch |
| 估值历史 | yfinance |
