#!/usr/bin/env python3
"""生成品牌会员投放数据 Dashboard (交互式 HTML)"""
import zipfile, re, json
from xml.etree import ElementTree as ET
from collections import defaultdict, Counter

def col_letter_to_index(col_str):
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

def parse_cell_value(cell_xml, strings):
    m = re.search(r'<v>(.*?)</v>', cell_xml)
    if not m:
        return None
    val = m.group(1)
    if 't="s"' in cell_xml or 't="inlineStr"' in cell_xml:
        try:
            idx = int(val)
            return strings[idx] if idx < len(strings) else val
        except:
            return val
    return val

print("Reading data...")
with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    # 列名映射
    col_names = strings[:26]
    COL_BRAND = 0
    COL_TARGET = 1     # 目标类型name: 27=拉新, 32=复购
    COL_DATE = 2
    COL_PLAN = 3
    COL_STORE = 4
    COL_CITY = 5
    COL_PLAN_NAME = 7
    COL_COST = 9
    COL_NEW_MEMBER = 10
    COL_CPA = 11
    COL_ROI = 12
    COL_TRADE_MEMBER = 13
    COL_CLICK = 14
    COL_IMPR = 15
    COL_REG_RATE = 16
    COL_CVR = 17
    COL_ORDERS = 18
    COL_CPM = 19
    COL_CPC = 20
    COL_CTR = 21
    COL_ORDER_PEOPLE = 22
    COL_GMV_ORIG = 23
    COL_GMV_PAID = 24

    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # 按日期聚合
    daily = defaultdict(lambda: {
        'cost': 0, 'new_members': 0, 'gmv': 0, 'orders': 0,
        'trade_members': 0, 'clicks': 0, 'impr': 0,
        'roi_sum': 0, 'roi_cnt': 0,
        'type_cost': defaultdict(float),
        'type_members': defaultdict(int),
        'city_cost': defaultdict(float),
        'city_members': defaultdict(int)
    })
    
    print(f"Processing {len(rows)-1} rows...")
    for idx, row_xml in enumerate(rows[1:]):
        if idx % 50000 == 0:
            print(f"  Row {idx}...")
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        row = {}
        for ref, content in cells:
            col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
            val = parse_cell_value(f'<c>{content}</c>', strings)
            row[col] = val
        
        date = row.get(COL_DATE)
        if not date:
            continue
        
        d = daily[date]
        
        cost = float(row.get(COL_COST) or 0)
        d['cost'] += cost
        
        nm = int(float(row.get(COL_NEW_MEMBER) or 0))
        d['new_members'] += nm
        
        gmv = float(row.get(COL_GMV_PAID) or 0)
        d['gmv'] += gmv
        
        orders = int(float(row.get(COL_ORDERS) or 0))
        d['orders'] += orders
        
        tm = int(float(row.get(COL_TRADE_MEMBER) or 0))
        d['trade_members'] += tm
        
        clicks = int(float(row.get(COL_CLICK) or 0))
        d['clicks'] += clicks
        
        impr = int(float(row.get(COL_IMPR) or 0))
        d['impr'] += impr
        
        roi = float(row.get(COL_ROI) or 0)
        if roi > 0:
            d['roi_sum'] += roi * cost  # 加权ROI
            d['roi_cnt'] += cost
        
        t = row.get(COL_TARGET)
        if t:
            is_acq = '27789310ba1cdb2fe93dc7eb23c367ea2c14b31e76dc0e8c58b2bc8d8af073ea'
            d['type_cost'][t] += cost
            d['type_members'][t] += nm
        
        city = row.get(COL_CITY)
        if city and city not in ('0', 'NULL', '', None):
            d['city_cost'][city] += cost
            d['city_members'][city] += nm
    
    # 准备图表数据
    sorted_dates = sorted(daily.keys())
    
    # 日维度数据
    dates = []
    cost_data = []
    member_data = []
    gmv_data = []
    order_data = []
    roi_data = []
    cpa_data = []
    impr_data = []
    click_data = []
    
    for d in sorted_dates:
        dd = daily[d]
        dates.append(d)
        cost_data.append(round(dd['cost'], 2))
        member_data.append(dd['new_members'])
        gmv_data.append(round(dd['gmv'], 2))
        order_data.append(dd['orders'])
        roi_data.append(round(dd['gmv'] / dd['cost'] if dd['cost'] > 0 else 0, 2))
        cpa_data.append(round(dd['cost'] / dd['new_members'] if dd['new_members'] > 0 else 0, 2))
        impr_data.append(dd['impr'])
        click_data.append(dd['clicks'])
    
    # 周维度聚合
    weeks = defaultdict(lambda: {'cost': 0, 'members': 0, 'gmv': 0})
    for d, dd in daily.items():
        import datetime
        dt = datetime.datetime.strptime(d, '%Y%m%d')
        week_str = dt.strftime('%Y-W%W')
        weeks[week_str]['cost'] += dd['cost']
        weeks[week_str]['members'] += dd['new_members']
        weeks[week_str]['gmv'] += dd['gmv']
    
    week_labels = sorted(weeks.keys())
    week_data = {
        'labels': week_labels,
        'cost': [round(weeks[w]['cost'], 2) for w in week_labels],
        'members': [weeks[w]['members'] for w in week_labels],
        'gmv': [round(weeks[w]['gmv'], 2) for w in week_labels]
    }
    
    # 城市TOP20
    all_city_cost = Counter()
    all_city_members = Counter()
    all_city_gmv = Counter()
    for d, dd in daily.items():
        for city, cost in dd['city_cost'].items():
            all_city_cost[city] += cost
        for city, m in dd['city_members'].items():
            all_city_members[city] += m
        # 按城市分摊GMV（粗略，用比例）
        for city in dd['city_cost']:
            ratio = dd['city_cost'][city] / dd['cost'] if dd['cost'] > 0 else 0
            all_city_gmv[city] += dd['gmv'] * ratio
    
    top_cities = [c for c, _ in all_city_cost.most_common(20)]
    top_cost = [round(all_city_cost[c], 2) for c in top_cities]
    top_members = [all_city_members[c] for c in top_cities]
    top_roi = [round(all_city_gmv[c] / all_city_cost[c] if all_city_cost[c] > 0 else 0, 2) for c in top_cities]
    
    # 目标类型对比
    type_acq = {'cost': 0, 'members': 0, 'gmv': 0}
    type_ret = {'cost': 0, 'members': 0, 'gmv': 0}
    for d, dd in daily.items():
        for t, c in dd['type_cost'].items():
            target = type_acq if t == '27' else type_ret
            target['cost'] += c
            target['members'] += dd['type_members'].get(t, 0)

print("Data processed. Generating HTML dashboard...")

# ====== 生成 HTML Dashboard ======
html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>霸王茶姬 · 品牌会员投放 Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family: -apple-system, "PingFang SC", "Helvetica Neue", sans-serif; background: #f5f7fa; color: #1a1a2e; padding: 20px; }}
.header {{ display:flex; justify-content:space-between; align-items:center; margin-bottom:24px; }}
.header h1 {{ font-size:24px; font-weight:700; color:#1a1a2e; }}
.header .sub {{ color:#6b7280; font-size:14px; }}
.filters {{ display:flex; gap:16px; margin-bottom:20px; flex-wrap:wrap; }}
.filters label {{ font-size:13px; font-weight:600; color:#374151; }}
.filters select, .filters input {{ padding:6px 12px; border:1px solid #d1d5db; border-radius:6px; font-size:13px; background:white; }}
.kpi-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(180px, 1fr)); gap:12px; margin-bottom:24px; }}
.kpi-card {{ background:white; border-radius:12px; padding:16px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
.kpi-card .label {{ font-size:12px; color:#6b7280; margin-bottom:4px; }}
.kpi-card .value {{ font-size:24px; font-weight:700; color:#1a1a2e; }}
.kpi-card .change {{ font-size:12px; color:#10b981; margin-top:2px; }}
.chart-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:24px; }}
.chart-full {{ grid-column: 1 / -1; }}
.chart-box {{ background:white; border-radius:12px; padding:12px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }}
.chart-box h3 {{ font-size:14px; font-weight:600; color:#374151; margin-bottom:8px; padding-left:4px; }}
@media (max-width: 900px) {{ .chart-grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>🫖 霸王茶姬 · 品牌会员投放 Dashboard</h1>
    <div class="sub">数据周期: {sorted_dates[0][:4]}-{sorted_dates[0][4:6]}-{sorted_dates[0][6:]} ~ {sorted_dates[-1][:4]}-{sorted_dates[-1][4:6]}-{sorted_dates[-1][6:]}</div>
  </div>
</div>

<div class="filters">
  <div><label>时间粒度</label><select id="timeGranularity" onchange="updateCharts()"><option value="daily">按日</option><option value="weekly">按周</option><option value="monthly">按月</option></select></div>
  <div><label>数据范围</label><select id="dateRange" onchange="updateCharts()"><option value="all">全部 (38天)</option><option value="last7">最近7天</option><option value="last14">最近14天</option><option value="last30">最近30天</option></select></div>
  <div><label>目标类型</label><select id="targetType" onchange="updateCharts()"><option value="all">全部</option><option value="acq">最大化新会员注册</option><option value="ret">最大化老会员转化</option></select></div>
  <div><label>城市</label><select id="cityFilter" onchange="updateCharts()"><option value="all">全部城市</option>{''.join(f'<option value="{c}">{c}</option>' for c in top_cities)}</select></div>
</div>

<div class="kpi-grid">
  <div class="kpi-card"><div class="label">总消耗</div><div class="value" id="kpi-cost">{sum(cost_data):,.2f}</div></div>
  <div class="kpi-card"><div class="label">新增会员</div><div class="value" id="kpi-members">{sum(member_data):,}</div></div>
  <div class="kpi-card"><div class="label">综合 CPA</div><div class="value" id="kpi-cpa">{sum(cost_data)/sum(member_data):.2f}</div></div>
  <div class="kpi-card"><div class="label">广告 ROI</div><div class="value" id="kpi-roi">{sum(gmv_data)/sum(cost_data):.2f}</div></div>
  <div class="kpi-card"><div class="label">实付 GMV</div><div class="value" id="kpi-gmv">{sum(gmv_data):,.2f}</div></div>
  <div class="kpi-card"><div class="label">订单数</div><div class="value" id="kpi-orders">{sum(order_data):,}</div></div>
</div>

<div class="chart-grid">
  <div class="chart-full chart-box">
    <h3>📈 每日会员增长 & 消耗趋势</h3>
    <div id="chart-trend" style="height:360px;"></div>
  </div>
  <div class="chart-box">
    <h3>📊 CPA & ROI 趋势</h3>
    <div id="chart-efficiency" style="height:320px;"></div>
  </div>
  <div class="chart-box">
    <h3>🔄 新增会员 vs 订单数</h3>
    <div id="chart-members-orders" style="height:320px;"></div>
  </div>
  <div class="chart-box">
    <h3>🏆 TOP20 城市 · 消耗排行</h3>
    <div id="chart-city-cost" style="height:360px;"></div>
  </div>
  <div class="chart-box">
    <h3>📌 城市效率矩阵 · 消耗 vs ROI</h3>
    <div id="chart-city-scatter" style="height:360px;"></div>
  </div>
  <div class="chart-box">
    <h3>📊 拉新 vs 复购 · 目标类型对比</h3>
    <div id="chart-target-compare" style="height:320px;"></div>
  </div>
</div>

<script>
// ====== RAW DATA ======
const rawData = {json.dumps({
    'dates': sorted_dates,
    'cost': cost_data,
    'members': member_data,
    'gmv': gmv_data,
    'orders': order_data,
    'roi': roi_data,
    'cpa': cpa_data,
    'impr': impr_data,
    'clicks': click_data
})};
const weeklyData = {json.dumps(week_data)};
const topCities = {json.dumps(top_cities)};
const topCost = {json.dumps(top_cost)};
const topMembers = {json.dumps(top_members)};
const topRoi = {json.dumps(top_roi)};

const allRawCost = [...rawData.cost];
const allRawMembers = [...rawData.members];
const allRawGmv = [...rawData.gmv];
const allRawOrders = [...rawData.orders];
const allRawRoi = [...rawData.roi];
const allRawCpa = [...rawData.cpa];

const colors = {{ primary: '#2563eb', green: '#10b981', red: '#ef4444', yellow: '#f59e0b', purple: '#8b5cf6' }};

function getFilteredIndices() {{
  const range = document.getElementById('dateRange').value;
  const n = rawData.dates.length;
  if (range === 'all') return {{start:0, end:n-1}};
  const days = {{'last7':7, 'last14':14, 'last30':30}}[range] || n;
  return {{start: Math.max(0, n - days), end: n - 1}};
}}

function updateKPI() {{  
  const {{start, end}} = getFilteredIndices();
  const c = rawData.cost.slice(start, end+1).reduce((a,b)=>a+b, 0);
  const m = rawData.members.slice(start, end+1).reduce((a,b)=>a+b, 0);
  const g = rawData.gmv.slice(start, end+1).reduce((a,b)=>a+b, 0);
  const o = rawData.orders.slice(start, end+1).reduce((a,b)=>a+b, 0);
  document.getElementById('kpi-cost').textContent = c.toLocaleString(undefined, {{minimumFractionDigits:2, maximumFractionDigits:2}});
  document.getElementById('kpi-members').textContent = m.toLocaleString();
  document.getElementById('kpi-cpa').textContent = m > 0 ? (c/m).toFixed(2) : '0.00';
  document.getElementById('kpi-roi').textContent = c > 0 ? (g/c).toFixed(2) : '0.00';
  document.getElementById('kpi-gmv').textContent = g.toLocaleString(undefined, {{minimumFractionDigits:2, maximumFractionDigits:2}});
  document.getElementById('kpi-orders').textContent = o.toLocaleString();
}}

function updateCharts() {{
  updateKPI();
  const {{start, end}} = getFilteredIndices();
  const granularity = document.getElementById('timeGranularity').value;
  
  let x, cost, members, gmv, orders, roi, cpa;
  
  if (granularity === 'weekly' && weeklyData.labels.length > 0) {{
    x = weeklyData.labels;
    cost = weeklyData.cost;
    members = weeklyData.members;
    gmv = weeklyData.gmv;
    // Estimate weekly roi / cpa
    roi = cost.map((c,i) => gmv[i]/c || 0);
    cpa = cost.map((c,i) => members[i] > 0 ? c/members[i] : 0);
  }} else {{
    x = rawData.dates.slice(start, end+1);
    cost = rawData.cost.slice(start, end+1);
    members = rawData.members.slice(start, end+1);
    gmv = rawData.gmv.slice(start, end+1);
    orders = rawData.orders.slice(start, end+1);
    roi = rawData.roi.slice(start, end+1);
    cpa = rawData.cpa.slice(start, end+1);
  }}
  
  // Trend chart: cost + members
  const trend1 = {{
    x: x, y: cost, type: 'scatter', mode: 'lines+markers', name: '消耗',
    line: {{color: colors.primary, width: 2}}, marker: {{size: 4}},
    yaxis: 'y', hovertemplate: '%{{x}}<br>消耗: ¥%{{y:,.2f}}<extra></extra>'
  }};
  const trend2 = {{
    x: x, y: members, type: 'scatter', mode: 'lines+markers', name: '新增会员',
    line: {{color: colors.green, width: 2}}, marker: {{size: 4}},
    yaxis: 'y2', hovertemplate: '%{{x}}<br>新增: %{{y:,}}人<extra></extra>'
  }};
  Plotly.newReact('chart-trend', {{
    data: [trend1, trend2],
    layout: {{
      margin: {{l:60, r:60, t:10, b:40}}, hovermode:'x unified',
      xaxis: {{type: 'category', tickangle: -45, tickfont: {{size:10}} }},
      yaxis: {{title: '消耗 (¥)', side: 'left', showgrid: true, gridcolor: '#f0f0f0'}},
      yaxis2: {{title: '新增会员', side: 'right', overlaying: 'y', showgrid: false}},
      legend: {{orientation: 'h', y: 1.08}},
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
  
  // Efficiency chart: CPA + ROI
  const eff1 = {{
    x: x, y: cpa, type: 'scatter', mode: 'lines+markers', name: 'CPA',
    line: {{color: colors.red, width: 2}}, marker: {{size: 4}},
    yaxis: 'y', hovertemplate: '%{{x}}<br>CPA: ¥%{{y:,.2f}}<extra></extra>'
  }};
  const eff2 = {{
    x: x, y: roi, type: 'scatter', mode: 'lines+markers', name: 'ROI',
    line: {{color: colors.yellow, width: 2}}, marker: {{size: 4}},
    yaxis: 'y2', hovertemplate: '%{{x}}<br>ROI: %{{y:.2f}}<extra></extra>'
  }};
  Plotly.newReact('chart-efficiency', {{
    data: [eff1, eff2],
    layout: {{
      margin: {{l:60, r:60, t:10, b:40}}, hovermode:'x unified',
      xaxis: {{type: 'category', tickangle: -45, tickfont: {{size:10}} }},
      yaxis: {{title: 'CPA (¥)', side: 'left', showgrid: true, gridcolor: '#f0f0f0'}},
      yaxis2: {{title: 'ROI', side: 'right', overlaying: 'y', showgrid: false, rangemode: 'tozero'}},
      legend: {{orientation: 'h', y: 1.08}},
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
  
  // Members vs Orders
  const mo1 = {{
    x: x, y: members, type: 'bar', name: '新增会员',
    marker: {{color: colors.green, opacity: 0.7}},
    yaxis: 'y', hovertemplate: '%{{x}}<br>新增: %{{y:,}}人<extra></extra>'
  }};
  const mo2 = {{
    x: x, y: orders, type: 'scatter', mode: 'lines+markers', name: '订单数',
    line: {{color: colors.purple, width: 2}}, marker: {{size: 4}},
    yaxis: 'y2', hovertemplate: '%{{x}}<br>订单: %{{y:,}}<extra></extra>'
  }};
  Plotly.newReact('chart-members-orders', {{
    data: [mo1, mo2],
    layout: {{
      margin: {{l:60, r:60, t:10, b:40}}, hovermode:'x unified',
      xaxis: {{type: 'category', tickangle: -45, tickfont: {{size:10}} }},
      yaxis: {{title: '新增会员', side: 'left', showgrid: true, gridcolor: '#f0f0f0'}},
      yaxis2: {{title: '订单数', side: 'right', overlaying: 'y', showgrid: false}},
      legend: {{orientation: 'h', y: 1.08}},
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
  
  // City cost bar
  const cityBar = {{
    x: topCities, y: topCost, type: 'bar',
    marker: {{
      color: topCost.map(c => c > topCost[4] ? colors.primary : '#93c5fd'),
      opacity: 0.85
    }},
    hovertemplate: '%{{x}}<br>消耗: ¥%{{y:,.2f}}<extra></extra>'
  }};
  Plotly.newReact('chart-city-cost', {{
    data: [cityBar],
    layout: {{
      margin: {{l:50, r:20, t:10, b:80}}, hovermode:'x unified',
      xaxis: {{type: 'category', tickangle: -45, tickfont: {{size:10}} }},
      yaxis: {{title: '消耗 (¥)', showgrid: true, gridcolor: '#f0f0f0'}},
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
  
  // City scatter: cost vs members vs roi
  const cityScatter = {{
    x: topCost, y: topRoi, text: topCities,
    type: 'scatter', mode: 'markers+text',
    marker: {{
      size: topMembers.map(m => Math.max(10, Math.min(60, m/2000))),
      color: topRoi, colorscale: 'RdYlGn', showscale: true,
      colorbar: {{title: 'ROI', len: 0.6}}
    }},
    textposition: 'top center', textfont: {{size: 10}},
    hovertemplate: '%{{text}}<br>消耗: ¥%{{x:,.2f}}<br>ROI: %{{y:.2f}}<br>新增: %{{marker.size}}k<extra></extra>'
  }};
  Plotly.newReact('chart-city-scatter', {{
    data: [cityScatter],
    layout: {{
      margin: {{l:60, r:40, t:10, b:40}},
      xaxis: {{title: '消耗 (¥)', type: 'log', showgrid: true, gridcolor: '#f0f0f0'}},
      yaxis: {{title: 'ROI', showgrid: true, gridcolor: '#f0f0f0', rangemode: 'tozero'}},
      shapes: [
        {{type: 'line', x0: 10000, y0: 2, x1: 300000, y1: 2, line: {{dash: 'dash', color: '#d1d5db', width: 1}}}}
      ],
      annotations: [
        {{x: 15000, y: 2.05, text: 'ROI=2 基准线', showarrow: false, font: {{size: 10, color: '#9ca3af'}}}}
      ],
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
  
  // Target type comparison (static for now)
  Plotly.newReact('chart-target-compare', {{
    data: [
      {{x: ['拉新·新增会员', '复购·新增会员'], y: [{json.dumps(type_acq['members'])}, {json.dumps(type_ret['members'])}], type: 'bar', name: '新增会员', marker: {{color: [colors.green, colors.primary]}}}},
      {{x: ['拉新·消耗', '复购·消耗'], y: [{json.dumps(round(type_acq['cost'],2))}, {json.dumps(round(type_ret['cost'],2))}], type: 'bar', name: '消耗', marker: {{color: [colors.primary, colors.purple]}}}}
    ],
    layout: {{
      barmode: 'group',
      margin: {{l:60, r:20, t:10, b:40}}, hovermode:'x',
      xaxis: {{type: 'category'}},
      yaxis: {{title: '值', showgrid: true, gridcolor: '#f0f0f0'}},
      legend: {{orientation: 'h', y: 1.08}},
      paper_bgcolor: 'white', plot_bgcolor: 'white'
    }}, config: {{responsive: true, displayModeBar: false}}
  }});
}}

// Initial render
updateCharts();
</script>
</body>
</html>'''

out_path = '/Users/edy/Desktop/brand_member_dashboard.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Dashboard saved to: {out_path}")
print("Done!")
