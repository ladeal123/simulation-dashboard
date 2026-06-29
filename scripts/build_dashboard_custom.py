"""定制仪表盘: 净值曲线+逐年收益+回撤+交易数+申万行业统计 (无交易明细表)"""
import json, csv, sys, openpyxl
from collections import Counter, defaultdict
sys.path.insert(0, '/workspace/skills/quant-backtest-lab/reference')
from render_dashboard import build_dashboard_data, render_dashboard

prefix = 'v48_final_d2d3'

# 读申万行业映射
print("加载行业数据...")
wb = openpyxl.load_workbook('/workspace/股票池_new.xlsx', data_only=True)
ws = wb['股票池']
code_to_industry = {}
for r in range(2, ws.max_row + 1):
    code = ws.cell(row=r, column=1).value
    ind = ws.cell(row=r, column=3).value
    if code and ind:
        code_to_industry[str(code).strip()] = str(ind).strip()
print(f"  行业映射: {len(code_to_industry)}只股票")

# 读原始JSON获取交易数据
with open('/workspace/V48_final_D2D3_result.json') as f:
    raw = json.load(f)
raw_trades = raw.get('trades', [])

# 逐年收益
equity = []
with open(f'{prefix}_equity.csv') as f:
    for row in csv.DictReader(f):
        equity.append(row)

# 逐年收益
yr_returns = defaultdict(list)
for e in equity:
    yr_returns[e['date'][:4]].append(float(e['value']))
yr_lines = []; yr_map = {}
for y in sorted(yr_returns.keys()):
    vals = yr_returns[y]
    ret = (vals[-1]/vals[0] - 1)*100
    yr_lines.append(f"{y}: {ret:+.2f}%")
    yr_map[y] = ret

# 收益分层
r22 = 1.0; r24 = 1.0
for y, ret in yr_map.items():
    if y >= '2022': r22 *= (1+ret/100)
    if y >= '2024': r24 *= (1+ret/100)

# 全部卖出原因 (从原始JSON)
reason_counter = Counter(t.get('卖出原因', '') for t in raw_trades)
reason_lines = [f"{k}: {v}笔" for k, v in reason_counter.most_common()]

# 逐年申万行业统计 (从原始JSON)
yearly_industry = {}
for t in raw_trades:
    y = t.get('日期', '')[:4] or t.get('卖出日期', '')[:4]
    code = t.get('代码', t.get('symbol', ''))
    ind = code_to_industry.get(code, '其他')
    if y not in yearly_industry:
        yearly_industry[y] = Counter()
    yearly_industry[y][ind] += 1

# 全期Top10行业 (从原始JSON)
all_ind_counter = Counter()
for t in raw_trades:
    code = t.get('代码', t.get('symbol', ''))
    ind = code_to_industry.get(code, '其他')
    all_ind_counter[ind] += 1
total_trades = len(raw_trades)
all_ind_lines = [f"{ind}: {cnt}笔({cnt/total_trades*100:.1f}%)"
                 for ind, cnt in all_ind_counter.most_common(10)]

# 逐年行业Top5
yearly_ind_lines = []
for y in sorted(yearly_industry.keys()):
    top5 = yearly_industry[y].most_common(5)
    yr_total = sum(yearly_industry[y].values())
    ind_strs = [f"{ind}({cnt/yr_total*100:.0f}%)" for ind, cnt in top5]
    yearly_ind_lines.append(f"{y}: {' '.join(ind_strs)}")

# 构建仪表盘
print("构建仪表盘...")
report_data = build_dashboard_data(
    equity_csv=f'{prefix}_equity.csv',
    summary_json=f'{prefix}_summary.json',
    language='zh', market='china_a',
    extra_modules=[
        {"type": "text", "tab": "overview", "title": "逐年收益",
         "text": "\n".join(yr_lines)},
        {"type": "text", "tab": "overview", "title": "分层收益",
         "text": f"2022至今: {(r22-1)*100:+.2f}%\n2024至今: {(r24-1)*100:+.2f}%"},
        {"type": "text", "tab": "overview", "title": "卖出原因",
         "text": "\n".join(reason_lines)},
        {"type": "text", "tab": "overview", "title": "全期申万行业Top10",
         "text": "\n".join(all_ind_lines)},
        {"type": "text", "tab": "overview", "title": "逐年交易行业分布(前5)",
         "text": "\n".join(yearly_ind_lines)},
    ],
)

# 移除交易明细表
report_data['modules'] = [m for m in report_data['modules'] if m.get('type') != 'trades_table']

render_dashboard(report_data, output_path='index_final.html',
    template_path='/workspace/skills/quant-backtest-lab/reference/dashboard_template.html')
print("Done!")
