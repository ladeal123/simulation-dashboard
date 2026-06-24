# 模拟盘 Dashboard - V46_v62_v20_final

## 策略说明
MACD(10,20,9) 金叉选股策略的模拟盘回测与可视化Dashboard。

### 策略要点
- 选股: BAR>0, 价格>MA144, 3日内金叉
- 评分: (DIF%×0.3 + BAR%×0.7) - 0.2×vol10
- 量比加分: ≥1.0/1.2/1.5 → +10%/+20%/+35%
- 执行: T日收盘信号, T+1开盘成交(含交易成本)
- 7年回测: +557.79%, 最大回撤19.43%

### 文件说明
| 路径 | 说明 |
|:---|:---|
| `backtest/` | 7年回测代码和结果 (final/v2/v3 三个版本) |
| `scripts/` | 模拟盘运行脚本 + Dashboard生成脚本 |
| `dashboard/` | 静态Dashboard网页 |
| `data/` | 每日数据文件(Wind导出的xlsx) |
| `.github/workflows/` | GitHub Actions 自动更新 |

### 本地运行
```bash
# 1. 安装依赖
pip install openpyxl

# 2. 生成Dashboard数据
python scripts/build_dashboard.py data/模拟盘_xxx.xlsx

# 3. 打开网页
# 将生成的 data.js 放到 dashboard/ 目录, 双击 index.html
```

### 每日更新
1. 从Wind导出最新数据到 `data/` 目录
2. 运行 `scripts/sim_V46_v62_v20_final.py` 生成模拟盘Excel
3. 运行 `scripts/build_dashboard.py` 生成 data.js
4. 推送到GitHub → GitHub Pages 自动更新
