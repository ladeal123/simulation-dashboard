"""模拟盘 Dashboard - Flask + Plotly (含行业+盈亏分析)"""
import json, math
from flask import Flask, render_template_string

app = Flask(__name__)

with open('/workspace/dashboard_data.json') as f:
    DATA = json.load(f)

HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>模拟盘 Dashboard - V46_v62_v20_final</title>
    <script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: -apple-system, 'PingFang SC', 'Microsoft YaHei', sans-serif; }
        body { background: #f0f2f5; color: #1a1a2e; }
        .header { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); padding: 28px 40px; position: sticky; top:0; z-index:100; }
        .header h1 { color: #fff; font-size: 24px; font-weight: 600; }
        .header p { color: #8899aa; font-size: 14px; margin-top: 6px; }
        .container { max-width: 1400px; margin: 0 auto; padding: 24px; }
        .nav-bar { display: flex; gap: 8px; margin-bottom: 24px; flex-wrap: wrap; }
        .nav-btn { padding: 8px 18px; border-radius: 20px; border: none; cursor: pointer; font-size: 13px; font-weight: 500; background: #e8ecf1; color: #555; }
        .nav-btn.active { background: #1a1a2e; color: #fff; }
        .section { display: none; } .section.active { display: block; }
        
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 24px; }
        .kpi-card { background: #fff; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
        .kpi-value { font-size: 28px; font-weight: 700; margin: 4px 0; }
        .kpi-label { font-size: 12px; color: #8899aa; }
        .kpi-green { color: #00c853; } .kpi-red { color: #ff5252; } .kpi-blue { color: #2196f3; } .kpi-orange { color: #ff9100; }
        
        .chart-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 24px; }
        .chart-card { background: #fff; border-radius: 12px; padding: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }
        .chart-card.full { grid-column: 1 / -1; }
        .chart-title { font-size: 14px; font-weight: 600; color: #333; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }
        .chart-wrap { width: 100%; height: 360px; }
        .chart-wrap.tall { height: 480px; } .chart-wrap.short { height: 300px; } .chart-wrap.mid { height: 380px; }
        
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        th { background: #f5f7fa; padding: 10px 12px; text-align: left; font-weight: 600; color: #555; border-bottom: 2px solid #e8ecf1; white-space: nowrap; }
        td { padding: 8px 12px; border-bottom: 1px solid #f0f2f5; }
        tr:hover td { background: #f8f9fb; }
        .tag-buy { color: #ff5252; font-weight: 600; } .tag-sell { color: #00c853; font-weight: 600; }
        .tag-win { color: #00c853; font-weight: 600; } .tag-lose { color: #ff5252; font-weight: 600; }
        .scroll-wrap { max-height: 400px; overflow-y: auto; border-radius: 8px; border: 1px solid #f0f2f5; }
        .scroll-wrap::-webkit-scrollbar { width: 6px; }
        .scroll-wrap::-webkit-scrollbar-thumb { background: #d0d5dd; border-radius: 3px; }
        .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 500; }
        .badge-ind { background: #e3f2fd; color: #1565c0; }
        
        @media (max-width: 900px) { .chart-grid { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="header">
        <h1>模拟盘 V46_v62_v20_final</h1>
        <p>{{ stats['回测起始日'] }} ~ {{ stats['回测结束日'] }} | {{ stats['回测天数'] }} 个交易日 | 共 {{ trades_total }} 笔交易</p>
    </div>
    <div class="container">
        <!-- 导航 -->
        <div class="nav-bar">
            <button class="nav-btn active" onclick="switchTab('tab-kpi')">概览</button>
            <button class="nav-btn" onclick="switchTab('tab-industry')">行业分析</button>
            <button class="nav-btn" onclick="switchTab('tab-pnl')">盈亏分析</button>
            <button class="nav-btn" onclick="switchTab('tab-trades')">交易记录</button>
        </div>

        <!-- ========== 概览 ========== -->
        <div class="section active" id="tab-kpi">
            <div class="kpi-grid">
                <div class="kpi-card"><div class="kpi-label">最终净值</div><div class="kpi-value kpi-green">{{ "%.4f"|format(stats['最终净值']) }}</div></div>
                <div class="kpi-card"><div class="kpi-label">总收益</div><div class="kpi-value kpi-green">{{ "+%.2f%%"|format(stats['总收益%']) }}</div></div>
                <div class="kpi-card"><div class="kpi-label">年化收益</div><div class="kpi-value kpi-green">{{ "+%.2f%%"|format(stats['年化收益%']) }}</div></div>
                <div class="kpi-card"><div class="kpi-label">最大回撤</div><div class="kpi-value kpi-red">{{ "%.2f%%"|format(stats['最大回撤%']) }}</div></div>
                <div class="kpi-card"><div class="kpi-label">胜率</div><div class="kpi-value kpi-blue">{{ "%.1f%%"|format(stats['胜率%']) }}</div></div>
                <div class="kpi-card"><div class="kpi-label">总交易</div><div class="kpi-value kpi-orange">{{ stats['总交易数'] }}</div></div>
                <div class="kpi-card"><div class="kpi-label">平均持仓</div><div class="kpi-value kpi-blue">{{ "%.1f"|format(stats['平均持仓天数']) }}天</div></div>
                <div class="kpi-card"><div class="kpi-label">交易成本</div><div class="kpi-value" style="color:#666;font-size:20px;">{{ "%.0f"|format(stats['交易成本合计']) }}</div></div>
            </div>

            <div class="chart-grid">
                <div class="chart-card full">
                    <div class="chart-title">净值走势</div>
                    <div class="chart-wrap tall" id="chart-nav"></div>
                </div>
            </div>
            <div class="chart-grid">
                <div class="chart-card">
                    <div class="chart-title">卖出原因分布</div>
                    <div class="chart-wrap short" id="chart-sell-reason"></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">每日持仓数量</div>
                    <div class="chart-wrap short" id="chart-positions"></div>
                </div>
            </div>
            
            <div class="chart-card full" style="margin-bottom:24px;">
                <div class="chart-title">当前持仓 ({{ holdings|length }}只)</div>
                <div class="scroll-wrap">
                    <table>
                        <thead><tr>
                            <th>代码</th><th>名称</th><th>行业</th><th>持仓数量</th><th>买入价格</th><th>当前价格</th><th>盈亏%</th><th>市值</th><th>入池天数</th>
                        </tr></thead>
                        <tbody>
                            {% for h in holdings %}
                            <tr>
                                <td><strong>{{ h['代码'] }}</strong></td>
                                <td>{{ h['名称'] }}</td>
                                <td><span class="badge badge-ind">{{ h['行业'] }}</span></td>
                                <td>{{ h['持仓数量'] }}</td>
                                <td>{{ h['买入价格'] }}</td>
                                <td>{{ h['当前价格'] }}</td>
                                <td style="color:{{ 'green' if h['盈亏%']|float >= 0 else 'red' }};">{{ h['盈亏%'] }}%</td>
                                <td>{{ h['市值'] }}</td>
                                <td>{{ h['入池天数'] }}</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- ========== 行业分析 ========== -->
        <div class="section" id="tab-industry">
            <div class="chart-grid">
                <div class="chart-card">
                    <div class="chart-title">持仓行业分布 (市值)</div>
                    <div class="chart-wrap mid" id="chart-ind-holding"></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">行业盈亏表现</div>
                    <div class="chart-wrap mid" id="chart-ind-pnl"></div>
                </div>
            </div>
            <div class="chart-card full" style="margin-bottom:24px;">
                <div class="chart-title">行业盈亏明细 (按总盈亏金额排序)</div>
                <div class="scroll-wrap">
                    <table>
                        <thead><tr>
                            <th>行业</th><th>交易次数</th><th>盈利次数</th><th>胜率</th><th>平均盈亏%</th><th>总盈亏金额</th>
                        </tr></thead>
                        <tbody>
                            {% for ind in industry_pnl %}
                            <tr>
                                <td><strong>{{ ind['industry'] }}</strong></td>
                                <td>{{ ind['count'] }}</td>
                                <td>{{ ind['wins'] }}</td>
                                <td>{{ ind['win_rate'] }}%</td>
                                <td style="color:{{ 'green' if ind['avg_pnl'] > 0 else 'red' }};">{{ ind['avg_pnl'] }}%</td>
                                <td style="color:{{ 'green' if ind['total_pnl_amt'] > 0 else 'red' }};">{{ "%.0f"|format(ind['total_pnl_amt']) }}</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- ========== 盈亏分析 ========== -->
        <div class="section" id="tab-pnl">
            <div class="chart-grid">
                <div class="chart-card">
                    <div class="chart-title">盈亏分布 (卖出交易)</div>
                    <div class="chart-wrap mid" id="chart-pnl-dist"></div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">盈亏金额TOP10</div>
                    <div class="chart-wrap mid" id="chart-pnl-top"></div>
                </div>
            </div>
            <div class="chart-grid">
                <div class="chart-card">
                    <div class="chart-title">最赚钱TOP10</div>
                    <div class="scroll-wrap" style="max-height:360px;">
                        <table>
                            <thead><tr><th>日期</th><th>代码</th><th>名称</th><th>行业</th><th>盈亏%</th><th>盈亏金额</th><th>持仓天</th></tr></thead>
                            <tbody>
                                {% for t in top_winners %}
                                <tr>
                                    <td>{{ t['日期'] }}</td><td><strong>{{ t['代码'] }}</strong></td><td>{{ t['名称'] }}</td>
                                    <td><span class="badge badge-ind">{{ t['行业'] }}</span></td>
                                    <td class="tag-win">{{ t['盈亏%'] }}%</td><td class="tag-win">{{ t['盈亏金额'] }}</td><td>{{ t['持仓天数'] }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
                <div class="chart-card">
                    <div class="chart-title">最亏钱TOP10</div>
                    <div class="scroll-wrap" style="max-height:360px;">
                        <table>
                            <thead><tr><th>日期</th><th>代码</th><th>名称</th><th>行业</th><th>盈亏%</th><th>盈亏金额</th><th>持仓天</th></tr></thead>
                            <tbody>
                                {% for t in top_losers %}
                                <tr>
                                    <td>{{ t['日期'] }}</td><td><strong>{{ t['代码'] }}</strong></td><td>{{ t['名称'] }}</td>
                                    <td><span class="badge badge-ind">{{ t['行业'] }}</span></td>
                                    <td class="tag-lose">{{ t['盈亏%'] }}%</td><td class="tag-lose">{{ t['盈亏金额'] }}</td><td>{{ t['持仓天数'] }}</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <!-- ========== 交易记录 ========== -->
        <div class="section" id="tab-trades">
            <div class="chart-card full" style="margin-bottom:24px;">
                <div class="chart-title">最近一周交易记录 ({{ week_trades|length }}条)</div>
                <div class="scroll-wrap" style="max-height:600px;">
                    <table>
                        <thead><tr>
                            <th>日期</th><th>代码</th><th>名称</th><th>行业</th><th>方向</th><th>价格</th><th>数量</th><th>金额</th><th>盈亏%</th><th>持仓天</th><th>原因</th>
                        </tr></thead>
                        <tbody>
                            {% for t in week_trades %}
                            <tr>
                                <td>{{ t['日期'] }}</td><td><strong>{{ t['代码'] }}</strong></td><td>{{ t['名称'] }}</td>
                                <td><span class="badge badge-ind">{{ t['行业'] }}</span></td>
                                <td class="{{ 'tag-sell' if t['买卖方向']=='卖出' else 'tag-buy' }}">{{ t['买卖方向'] }}</td>
                                <td>{{ t['价格'] }}</td><td>{{ t['数量'] }}</td><td>{{ t['金额'] }}</td>
                                <td style="color:{{ 'green' if t['盈亏%']|float > 0 else ('red' if t['盈亏%']|float < 0 else '#999') }};">{{ t['盈亏%'] }}%</td>
                                <td>{{ t['持仓天数'] }}</td><td>{{ t['卖出原因'] }}</td>
                            </tr>
                            {% endfor %}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
    function switchTab(tab) {
        document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
        document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
        document.getElementById(tab).classList.add('active');
        event.target.classList.add('active');
    }

    var navData = {{ nav_json|safe }};
    var dates = navData.map(d => d.date);
    var navs = navData.map(d => d.nav);
    var peak = navs[0], dd = navs.map(n => { var d = (peak - n) / peak * 100; if (n > peak) peak = n; return -Math.max(0, d); });

    Plotly.newPlot('chart-nav', [
        { x: dates, y: navs, type: 'scatter', mode: 'lines', name: '净值', line: { color: '#2196f3', width: 2 }, yaxis: 'y', fill: 'tozeroy', fillcolor: 'rgba(33,150,243,0.08)' },
        { x: dates, y: dd, type: 'scatter', mode: 'lines', name: '回撤%', line: { color: '#ff5252', width: 1.5, dash: 'dot' }, yaxis: 'y2', fill: 'tozeroy', fillcolor: 'rgba(255,82,82,0.08)' }
    ], {
        margin: { t: 10, r: 60, b: 40, l: 60 }, hovermode: 'x unified',
        xaxis: { showgrid: true, gridcolor: '#f0f2f5' },
        yaxis: { title: '净值', side: 'left', showgrid: true, gridcolor: '#f0f2f5', tickformat: '.4f' },
        yaxis2: { title: '回撤%', side: 'right', overlaying: 'y', showgrid: false, tickformat: '.1f', range: [Math.min(...dd)*1.15, 2] },
        legend: { orientation: 'h', y: 1.08 }, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)'
    }, {responsive: true, displaylogo: false});

    // 卖出原因
    var sellStats = {{ sell_stats_json|safe }};
    Plotly.newPlot('chart-sell-reason', [{
        labels: sellStats.map(s => s[0]), values: sellStats.map(s => s[1]),
        type: 'pie', hole: 0.45,
        marker: { colors: ['#ff5252','#ff9100','#00c853','#2196f3','#9c27b0','#e91e63'] },
        textinfo: 'label+percent', textposition: 'outside', hoverinfo: 'label+value'
    }], { margin: { t: 10, b: 10, l: 10, r: 10 }, showlegend: false, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});

    // 持仓数量
    Plotly.newPlot('chart-positions', [{
        x: dates, y: navData.map(d => d.positions), type: 'bar', marker: { color: '#4caf50', opacity: 0.7 }
    }], { margin: { t: 10, r: 10, b: 40, l: 50 }, hovermode: 'x', xaxis: { showgrid: false }, yaxis: { title: '持仓数量', dtick: 10 }, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});

    // 行业持仓分布
    var indData = {{ ind_holding_json|safe }};
    Plotly.newPlot('chart-ind-holding', [{
        labels: indData.map(d => d.industry), values: indData.map(d => d.value),
        type: 'pie', hole: 0.45,
        textinfo: 'label+percent', textposition: 'outside', hoverinfo: 'label+value+percent'
    }], { margin: { t: 10, b: 10, l: 10, r: 10 }, showlegend: false, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});

    // 行业盈亏
    var indPnl = {{ ind_pnl_json|safe }};
    Plotly.newPlot('chart-ind-pnl', [
        { x: indPnl.map(d => d.industry), y: indPnl.map(d => d.avg_pnl), type: 'bar', name: '平均盈亏%', marker: { color: indPnl.map(d => d.avg_pnl > 0 ? '#00c853' : '#ff5252') } }
    ], { margin: { t: 10, r: 10, b: 80, l: 50 }, xaxis: { tickangle: -45, showgrid: false }, yaxis: { title: '平均盈亏%' }, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});

    // 盈亏分布
    var pnlDist = {{ pnl_dist_json|safe }};
    var pnlColors = {'<-10%':'#d32f2f','-10%~-5%':'#ff5252','-5%~0%':'#ffab91','0%~5%':'#a5d6a7','5%~10%':'#66bb6a','>10%':'#2e7d32'};
    Plotly.newPlot('chart-pnl-dist', [{
        x: pnlDist.map(d => d.range), y: pnlDist.map(d => d.count),
        type: 'bar', marker: { color: pnlDist.map(d => pnlColors[d.range] || '#999') }
    }], { margin: { t: 10, r: 10, b: 40, l: 50 }, xaxis: { title: '盈亏区间' }, yaxis: { title: '交易次数' }, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});

    // 盈亏TOP10
    var topWinners = {{ top_winners_json|safe }};
    var topLosers = {{ top_losers_json|safe }};
    var allTop = [...topWinners, ...topLosers];
    Plotly.newPlot('chart-pnl-top', [{
        x: allTop.map(d => d['代码'] + ' ' + d['名称']),
        y: allTop.map(d => d['盈亏%值']),
        type: 'bar',
        marker: { color: allTop.map(d => d['盈亏%值'] > 0 ? '#00c853' : '#ff5252') },
        text: allTop.map(d => d['盈亏%'] + '%'),
        textposition: 'outside'
    }], { margin: { t: 10, r: 10, b: 100, l: 50 }, xaxis: { tickangle: -45, showgrid: false }, yaxis: { title: '盈亏%' }, plot_bgcolor: 'rgba(0,0,0,0)', paper_bgcolor: 'rgba(0,0,0,0)' }, {responsive: true, displaylogo: false});
    </script>
</body>
</html>
'''

@app.route('/')
def dashboard():
    stats = DATA['stats']
    # 卖出原因
    sell_reasons = {}
    for k, v in stats.items():
        if '卖出原因' in k:
            reason = k.replace('卖出原因:', '')
            sell_reasons[reason] = v
    
    return render_template_string(HTML_TEMPLATE,
        stats=stats,
        nav_json=json.dumps(DATA['nav'], ensure_ascii=False),
        sell_stats_json=json.dumps(sorted(sell_reasons.items(), key=lambda x: -x[1]), ensure_ascii=False),
        holdings=DATA['holdings'],
        week_trades=DATA['week_trades'],
        industry_analysis=DATA['industry_analysis'],
        industry_pnl=DATA['industry_pnl'],
        top_winners=DATA['top_winners'],
        top_losers=DATA['top_losers'],
        ind_holding_json=json.dumps(DATA['industry_analysis'], ensure_ascii=False),
        ind_pnl_json=json.dumps(DATA['industry_pnl'], ensure_ascii=False),
        pnl_dist_json=json.dumps(DATA['pnl_distribution'], ensure_ascii=False),
        top_winners_json=json.dumps(DATA['top_winners'], ensure_ascii=False),
        top_losers_json=json.dumps(DATA['top_losers'], ensure_ascii=False),
        trades_total=DATA['trades_total'],
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8899, debug=False)
