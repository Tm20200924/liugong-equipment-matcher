# LiuGong Equipment Matcher v19

柳工设备智能匹配引擎 — 客户询盘 → 产品匹配 → DAP满洲里到岸价计算 → 交叉验证

## 功能

1. **智能匹配**: 上传Excel/Word询盘文件，自动识别设备类型、吨位、马力需求
2. **竞品对标**: 自动识别竞品型号（XCMG、SANY、Komatsu、CAT等），联网搜索参数
3. **DAP价格计算**: 完整的DAP满洲里到岸价（RUB），含关税、报废税、增值税
4. **交叉验证**: 联网搜索 + 本地数据库双重校验，标记验证等级（dual/single/conflict）
5. **报告导出**: Excel格式完整报告

## 快速开始

### 本地运行

`ash
pip install -r requirements.txt
streamlit run app.py
`

或双击 一键运行.bat

### 文件结构

`
excel/
├── app.py                  # Streamlit Web界面
├── engine/
│   ├── engine_v19.py       # 核心匹配引擎
│   ├── product_db_full.json # 产品数据库 (不提交git)
│   ├── competitor_db.json   # 竞品数据库 (不提交git)
│   └── config.json          # 汇率/税率配置 (不提交git)
├── inquiries/              # 询盘文件 (不提交git)
├── outputs/                # 输出报告
└── .streamlit/
    └── secrets.toml        # Google Drive配置 (不提交git)
`

## 配置说明

### 汇率与税率

通过 Streamlit 侧边栏实时调整：
- 汇率 (CNY→RUB)
- 关税 %
- 增值税 %
- 报关手续费 %
- 仓储费/海关费/代理费 (RUB)

### Google Drive 数据加载 (可选)

在 .streamlit/secrets.toml 中配置：
`	oml
[gdrive]
product_db_url = "https://drive.google.com/uc?id=YOUR_FILE_ID"
competitor_db_url = "https://drive.google.com/uc?id=YOUR_FILE_ID"
scrap_tax_url = "https://drive.google.com/uc?id=YOUR_FILE_ID"
`

## DAP计算公式

`
DAP满洲里(RUB) = (DAP CNY + 关税CNY + 报关手续费CNY + 卢布固定费用折算CNY) × (1 + 增值税率) × 汇率

其中:
- 关税CNY = DAP CNY × 关税%
- 报关手续费CNY = DAP CNY × 报关手续费%
- 卢布固定费用 = 报废税 + 仓储费 + 海关费 + 代理费
- 卢布固定费用折算CNY = 卢布固定费用 ÷ 汇率
`

## 验证等级

| 等级 | 含义 |
|------|------|
| dual | 联网数据与本地数据库一致 |
| single | 仅一方有数据 |
| close | 双方数据接近（偏差<15%） |
| conflict | 双方数据冲突（偏差>15%），需人工确认 |
| none | 无验证数据 |
