#!/usr/bin/env python3
"""从数据抽取结果生成改造版HTML看板"""
import json, subprocess

# Get data
result = subprocess.run(['python3', '/Users/edy/Desktop/extract_data_v2.py'], capture_output=True, text=True, timeout=300)
out = json.loads(result.stdout)

D = out
dates_json = json.dumps(D['dates'])

html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>霸王茶姬 · 品牌会员投放 Dashboard V2</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,"PingFang SC","Helvetica Neue",sans-serif;background:#f5f7fa;color:#1a1a2e;padding:20px;max-width:1400px;margin:0 auto}}
.header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:12px}}
.header h1{{font-size:22px;font-weight:700;color:#1a1a2e}}
.header .sub{{color:#6b7280;font-size:14px}}
.filters{{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:flex-end}}
.filters .fg{{display:flex;flex-direction:column}}
.filters label{{font-size:11px;font-weight:600;color:#374151;margin-bottom:2px}}
.filters select,.filters input{{padding:5px 9px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;background:#fff}}
.filters .btn{{padding:6px 16px;border:none;border-radius:6px;font-size:13px;cursor:pointer;background:#2563eb;color:#fff;font-weight:600}}
.filters .btn:hover{{background:#1d4ed8}}
.dimension-tabs{{display:flex;gap:0;margin-bottom:14px}}
.dimension-tabs .tab{{padding:8px 18px;border:1px solid #d1d5db;cursor:pointer;font-size:13px;font-weight:500;background:#fff;color:#374151}}
.dimension-tabs .tab:first-child{{border-radius:6px 0 0 6px}}
.dimension-tabs .tab:last-child{{border-radius:0 6px 6px 0}}
.dimension-tabs .tab.active{{background:#2563eb;color:#fff;border-color:#2563eb}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:18px}}
.kpi-card{{background:#fff;border-radius:10px;padding:12px 14px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}}
.kpi-card .l{{font-size:11px;color:#6b7280;margin-bottom:1px}}
.kpi-card .v{{font-size:19px;font-weight:700;color:#1a1a2e}}
.kpi-card .s{{font-size:11px;color:#9ca3af;margin-top:1px}}
.chart-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:20px}}
.chart-full{{grid-column:1/-1}}
.chart-box{{background:#fff;border-radius:10px;padding:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06);position:relative}}
.chart-box h3{{font-size:13px;font-weight:600;color:#374151;margin-bottom:4px;padding-left:2px}}
.insight{{font-size:11px;color:#6b7280;padding:4px 6px;background:#f0f7ff;border-radius:4px;margin-bottom:6px;border-left:3px solid #2563eb}}
@media(max-width:900px){{.chart-grid{{grid-template-columns:1fr}}}}
</style>
</head>
<body>
<div class="header">
  <div><h1>🥖 霸王茶姬 · 品牌会员投放 Dashboard V2</h1><div class="sub">2026-05-01 ~ 2026-06-07（38天）| 拉新+复购</div></div>
</div>

<div class="filters">
  <div class="fg"><label>维度切换</label><select id="dimension"><option value="province">省份</option><option value="city">城市</option></select></div>
  <div class="fg"><label>时间粒度</label><select id="grain"><option value="day">按日</option><option value="week">按周</option><option value="month">按月</option></select></div>
  <div class="fg"><label>开始</label><input type="date" id="ds"></div>
  <div class="fg"><label>结束</label><input type="date" id="de"></div>
  <div class="fg"><label>目标类型</label><select id="target"><option value="all">全部</option><option value="acq">新会员注册(拉新)</option><option value="ret">老会员转化(复购)</option></select></div>
  <div class="fg"><label>区域</label><select id="region"><option value="all">全部区域</option></select></div>
  <button class="btn" onclick="applyFilter()">🔄 更新筛选</button>
</div>

<div class="kpi-grid" id="kpiGrid"></div>

<div class="chart-grid">
  <div class="chart-full chart-box">
    <h3>📈 消耗 & 新增会员趋势（按日）</h3>
    <div id="insight1" class="insight">加载中...</div>
    <div id="c1" style="height:320px;"></div>
  </div>
  <div class="chart-box">
    <h3>💰 CAC & ROI 趋势</h3>
    <div id="insight2" class="insight">加载中...</div>
    <div id="c2" style="height:300px;"></div>
  </div>
  <div class="chart-box">
    <h3>🔄 拉新 vs 复购 · 按日趋势</h3>
    <div id="insight6" class="insight">加载中...</div>
    <div id="c6" style="height:300px;"></div>
  </div>
  <div class="chart-box">
    <h3>🏆 TOP15 区域消耗排行</h3>
    <div id="insight4" class="insight">加载中...</div>
    <div id="c4" style="height:340px;"></div>
  </div>
  <div class="chart-box">
    <h3>📌 区域效率矩阵（X=消耗, Y=新增会员, 气泡=ROI）</h3>
    <div id="insight5" class="insight">加载中...</div>
    <div id="c5" style="height:340px;"></div>
  </div>
  <div class="chart-box">
    <h3>📊 GMV 贡献排行</h3>
    <div id="insight3" class="insight">加载中...</div>
    <div id="c3" style="height:340px;"></div>
  </div>
</div>

<script>
const D = {json.dumps(D, ensure_ascii=False)};
const C = {{prim:'#2563eb',green:'#10b981',red:'#ef4444',yellow:'#f59e0b',purple:'#8b5cf6',orange:'#f97316'}};
let curDim = 'province';
let filteredDates = D.dates.slice();

function fmtDate(d) {{return d.slice(0,4)+'-'+d.slice(4,6)+'-'+d.slice(6,8);}}
function parseDate(d) {{return new Date(+d.slice(0,4),+d.slice(4,6)-1,+d.slice(6,8));}}

// Init date filters
document.getElementById('ds').value = fmtDate(D.dates[0]);
document.getElementById('de').value = fmtDate(D.dates[D.dates.length-1]);

// Init region select
const regions = curDim === 'province' ? D.provinces : D.cities;
const sel = document.getElementById('region');
sel.innerHTML = '<option value="all">全部区域</option>';
regions.forEach(r => {{var o=document.createElement('option');o.value=r;o.textContent=r;sel.appendChild(o)}});

// Dim tab switch
document.getElementById('dimension').addEventListener('change', function() {{
  curDim = this.value;
  const regions = curDim === 'province' ? D.provinces : D.cities;
  const sel = document.getElementById('region');
  sel.innerHTML = '<option value="all">全部区域</option>';
  regions.forEach(r => {{var o=document.createElement('option');o.value=r;o.textContent=r;sel.appendChild(o)}});
  applyFilter();
}});

function applyFilter() {{
  const ds = document.getElementById('ds').value.replace(/-/g,'');
  const de = document.getElementById('de').value.replace(/-/g,'');
  const target = document.getElementById('target').value;
  const region = document.getElementById('region').value;
  
  filteredDates = D.dates.filter(d => d >= ds && d <= de);
  
  updateKPIs(target);
  drawCharts(target, region);
}}

function updateKPIs(target) {{
  let cost, members, gmv, roi, cac;
  const idx = filteredDates.map(d => D.dates.indexOf(d));
  
  if (target === 'acq') {{
    cost = idx.reduce((s,i) => s + D.acq_cost[i], 0);
    members = idx.reduce((s,i) => s + D.acq_members[i], 0);
    gmv = idx.reduce((s,i) => s + D.acq_gmv[i], 0);
  }} else if (target === 'ret') {{
    cost = idx.reduce((s,i) => s + D.ret_cost[i], 0);
    members = idx.reduce((s,i) => s + D.ret_members[i], 0);
    gmv = 0;
  }} else {{
    cost = idx.reduce((s,i) => s + D.acq_cost[i] + D.ret_cost[i], 0);
    members = idx.reduce((s,i) => s + D.acq_members[i] + D.ret_members[i], 0);
    gmv = idx.reduce((s,i) => s + D.acq_gmv[i], 0);
  }}
  roi = gmv / cost || 0;
  cac = cost / members || 0;
  
  const days = filteredDates.length || 1;
  document.getElementById('kpiGrid').innerHTML =
    '<div class=kpi-card><div class=l>总消耗(CA)</div><div class=v>¥'+(cost/10000).toFixed(1)+'万</div><div class=s>日均¥'+(cost/days/10000).toFixed(2)+'万</div></div>' +
    '<div class=kpi-card><div class=l>注册会员(CAC)</div><div class=v>'+members.toLocaleString()+'人</div><div class=s>CAC ¥'+cac.toFixed(2)+'</div></div>' +
    '<div class=kpi-card><div class=l>总实付GMV</div><div class=v>¥'+(gmv/10000).toFixed(1)+'万</div><div class=s>ROI '+roi.toFixed(2)+'</div></div>' +
    '<div class=kpi-card><div class=l>日均新增会员</div><div class=v>'+(members/days).toFixed(1)+'</div><div class=s>38天共'+days+'天</div></div>';
}}

function getDimData(target, region) {{
  let names, costs, members, gmv, roi, cac;
  if (curDim === 'province') {{
    names = D.provinces;
    costs = D.prov_cost;
    members = D.prov_members;
    gmv = D.prov_gmv;
    roi = D.prov_roi;
    cac = D.prov_cac;
  }} else {{
    names = D.cities;
    costs = D.city_cost;
    members = D.city_members;
    gmv = D.city_gmv;
    roi = D.city_roi;
    cac = D.city_cac;
  }}
  
  if (region !== 'all') {{
    const ri = names.indexOf(region);
    if (ri >= 0) {{
      names = [names[ri]]; costs = [costs[ri]]; members = [members[ri]];
      gmv = [gmv[ri]]; roi = [roi[ri]]; cac = [cac[ri]];
    }}
  }}
  return {{names, costs, members, gmv, roi, cac}};
}}

function drawCharts(target, region) {{
  const idx = filteredDates.map(d => D.dates.indexOf(d));
  
  // ---- Chart 1: Trend ----
  let t1_cost, t1_members, t1_label;
  if (target === 'acq') {{
    t1_cost = idx.map(i => D.acq_cost[i]);
    t1_members = idx.map(i => D.acq_members[i]);
    t1_label = '拉新';
  }} else if (target === 'ret') {{
    t1_cost = idx.map(i => D.ret_cost[i]);
    t1_members = idx.map(i => D.ret_members[i]);
    t1_label = '复购';
  }} else {{
    t1_cost = idx.map(i => D.acq_cost[i] + D.ret_cost[i]);
    t1_members = idx.map(i => D.acq_members[i] + D.ret_members[i]);
    t1_label = '全部';
  }}
  
  Plotly.newPlot('c1', [
    {{x: filteredDates, y: t1_cost, type:'scatter', mode:'lines+markers', name:'消耗(CA)',
      line:{{color:C.prim,width:2}}, marker:{{size:5}},
      yaxis:'y', hovertemplate:'%{{x|%Y-%m-%d}}<br>消耗: ¥%{{y:,.2f}}<extra></extra>'}},
    {{x: filteredDates, y: t1_members, type:'scatter', mode:'lines+markers', name:'新增会员',
      line:{{color:C.green,width:2}}, marker:{{size:5}},
      yaxis:'y2', hovertemplate:'%{{x|%Y-%m-%d}}<br>新增: %{{y:,}}人<extra></extra>'}}
  ], {{
    margin:{{l:60,r:60,t:10,b:50}}, hovermode:'x unified',
    xaxis:{{tickangle:-45,tickfont:{{size:10}},tickformat:'%m/%d',
           tickmode:'linear',dtick:3}},
    yaxis:{{title:'消耗(¥)',side:'left',showgrid:true,gridcolor:'#f0f0f0'}},
    yaxis2:{{title:'新增会员',side:'right',overlaying:'y',showgrid:false}},
    legend:{{orientation:'h',y:1.08}},
    paper_bgcolor:'white',plot_bgcolor:'white',
    shapes:[{{type:'line',x0:filteredDates[0],x1:filteredDates[filteredDates.length-1],
             y0:0,y1:0,line:{{color:'transparent'}}}}]
  }}, {{responsive:true,displayModeBar:false,toImageButtonOptions:{{format:'png'}}}});
  
  // Insight 1
  const avgCost = (t1_cost.reduce((a,b)=>a+b,0)/t1_cost.length).toFixed(0);
  const maxCostIdx = t1_cost.indexOf(Math.max(...t1_cost));
  const maxMemberIdx = t1_members.indexOf(Math.max(...t1_members));
  document.getElementById('insight1').textContent =
    `💡 日均消耗¥${avgCost} | 消耗峰值 ${fmtDate(filteredDates[maxCostIdx])} (¥${(t1_cost[maxCostIdx]/10000).toFixed(1)}万) | 会员峰值 ${fmtDate(filteredDates[maxMemberIdx])} (${t1_members[maxMemberIdx]}人)`;
  
  // ---- Chart 2: CAC & ROI Trend ----
  let t2_cac, t2_roi;
  if (target === 'acq') {{
    t2_cac = idx.map(i => D.acq_cost[i] / (D.acq_members[i]||1));
    t2_roi = idx.map(i => D.acq_gmv[i] / (D.acq_cost[i]||1));
  }} else if (target === 'ret') {{
    t2_cac = idx.map(i => D.ret_cost[i] / (D.ret_members[i]||1));
    t2_roi = idx.map(() => 0);
  }} else {{
    t2_cac = idx.map(i => (D.acq_cost[i]+D.ret_cost[i]) / ((D.acq_members[i]+D.ret_members[i])||1));
    t2_roi = idx.map(i => D.acq_gmv[i] / ((D.acq_cost[i]+D.ret_cost[i])||1));
  }}
  
  Plotly.newPlot('c2', [
    {{x: filteredDates, y: t2_cac, type:'scatter', mode:'lines+markers', name:'CAC(注册成本)',
      line:{{color:C.red,width:2}}, marker:{{size:4}},
      yaxis:'y', hovertemplate:'%{{x|%Y-%m-%d}}<br>CAC: ¥%{{y:,.2f}}<extra></extra>'}},
    {{x: filteredDates, y: t2_roi, type:'scatter', mode:'lines+markers', name:'ROI',
      line:{{color:C.yellow,width:2}}, marker:{{size:4}},
      yaxis:'y2', hovertemplate:'%{{x|%Y-%m-%d}}<br>ROI: %{{y:.2f}}<extra></extra>'}}
  ], {{
    margin:{{l:50,r:50,t:10,b:50}}, hovermode:'x unified',
    xaxis:{{tickangle:-45,tickfont:{{size:10}},tickformat:'%m/%d',tickmode:'linear',dtick:3}},
    yaxis:{{title:'CAC(¥)',side:'left',showgrid:true,gridcolor:'#f0f0f0'}},
    yaxis2:{{title:'ROI',side:'right',overlaying:'y',showgrid:false,rangemode:'tozero'}},
    legend:{{orientation:'h',y:1.08}},
    paper_bgcolor:'white',plot_bgcolor:'white'
  }}, {{responsive:true,displayModeBar:false}});
  
  const avgCac = (t2_cac.filter(v=>v<1000).reduce((a,b)=>a+b,0)/t2_cac.filter(v=>v<1000).length||0).toFixed(1);
  const avgRoi = (t2_roi.reduce((a,b)=>a+b,0)/t2_roi.length).toFixed(2);
  document.getElementById('insight2').textContent =
    `💡 日均CAC ¥${avgCac} | 日均ROI ${avgRoi} | CAC波动${(Math.max(...t2_cac.filter(v=>v<1000))-Math.min(...t2_cac.filter(v=>v<1000))).toFixed(0)}`;
  
  // ---- Chart 6: 拉新 vs 复购 按日趋势 ----
  Plotly.newPlot('c6', [
    {{x: filteredDates, y: idx.map(i => D.acq_members[i]), type:'bar', name:'拉新·新增会员',
      marker:{{color:C.green,opacity:0.8}},
      yaxis:'y', hovertemplate:'%{{x|%Y-%m-%d}}<br>拉新: %{{y:,}}人<extra></extra>'}},
    {{x: filteredDates, y: idx.map(i => D.ret_members[i]), type:'bar', name:'复购·新增会员',
      marker:{{color:C.prim,opacity:0.8}},
      yaxis:'y', hovertemplate:'%{{x|%Y-%m-%d}}<br>复购: %{{y:,}}人<extra></extra>'}},
    {{x: filteredDates, y: idx.map(i => D.acq_cost[i]/10000), type:'scatter', mode:'lines+markers',
      name:'拉新消耗(¥万)', line:{{color:C.green,width:1.5,dash:'dot'}}, marker:{{size:3}},
      yaxis:'y2', hovertemplate:'%{{x|%Y-%m-%d}}<br>拉新消耗: ¥%{{y:.2f}}万<extra></extra>'}},
    {{x: filteredDates, y: idx.map(i => D.ret_cost[i]/10000), type:'scatter', mode:'lines+markers',
      name:'复购消耗(¥万)', line:{{color:C.prim,width:1.5,dash:'dot'}}, marker:{{size:3}},
      yaxis:'y2', hovertemplate:'%{{x|%Y-%m-%d}}<br>复购消耗: ¥%{{y:.2f}}万<extra></extra>'}}
  ], {{
    barmode:'stack', margin:{{l:50,r:60,t:10,b:50}}, hovermode:'x unified',
    xaxis:{{tickangle:-45,tickfont:{{size:10}},tickformat:'%m/%d',tickmode:'linear',dtick:3}},
    yaxis:{{title:'新增会员',side:'left',showgrid:true,gridcolor:'#f0f0f0'}},
    yaxis2:{{title:'消耗(¥万)',side:'right',overlaying:'y',showgrid:false}},
    legend:{{orientation:'h',y:1.08}},
    paper_bgcolor:'white',plot_bgcolor:'white'
  }}, {{responsive:true,displayModeBar:false}});
  
  const acqTotal = idx.reduce((s,i)=>s+D.acq_members[i],0);
  const retTotal = idx.reduce((s,i)=>s+D.ret_members[i],0);
  document.getElementById('insight6').textContent =
    `💡 拉新增会员 ${acqTotal.toLocaleString()}人 | 复购新增 ${retTotal}人 | 拉新占比 ${(acqTotal/(acqTotal+retTotal||1)*100).toFixed(1)}%`;
  
  // ---- Chart 4: TOP15 区域消耗排行 ----
  const dd = getDimData(target, region);
  const topN = dd.names.slice(0,15).reverse();
  const topCost = dd.costs.slice(0,15).reverse();
  
  Plotly.newPlot('c4', [
    {{x: topCost, y: topN, type:'bar', orientation:'h',
      marker:{{color: topCost.map(c => c > topCost[Math.floor(topCost.length/2)] ? C.prim : '#93c5fd'), opacity:0.85}},
      hovertemplate:'%{{y}}<br>消耗: ¥%{{x:,.2f}}<extra></extra>'}}
  ], {{
    margin:{{l:70,r:20,t:5,b:40}}, hovermode:'y unified',
    xaxis:{{title:'消耗(¥)', showgrid:true, gridcolor:'#f0f0f0'}},
    yaxis:{{tickfont:{{size:11}}, automargin:true}},
    paper_bgcolor:'white', plot_bgcolor:'white'
  }}, {{responsive:true,displayModeBar:false}});
  
  document.getElementById('insight4').textContent =
    `💡 TOP1: ${dd.names[0]} (¥${(dd.costs[0]/10000).toFixed(1)}万) | TOP3合计: ¥${((dd.costs[0]+dd.costs[1]+dd.costs[2])/10000).toFixed(1)}万 | 占比 ${((dd.costs[0]+dd.costs[1]+dd.costs[2])/dd.costs.reduce((a,b)=>a+b,0)*100).toFixed(0)}%`;
  
  // ---- Chart 5: 区域效率气泡图（X=消耗, Y=新增会员, 气泡=ROI, 颜色=CAC） ----
  const bd = getDimData(target, region);
  const bubbleData = bd.names.map((n,i) => ({{
    name: n, cost: bd.costs[i], members: bd.members[i],
    roi: bd.roi[i], cac: bd.cac[i]
  }})).filter(d => d.cost > 1000 && d.members > 0);
  
  if (bubbleData.length > 0) {{
    Plotly.newPlot('c5', [{
      x: bubbleData.map(d => d.cost),
      y: bubbleData.map(d => d.members),
      text: bubbleData.map(d => d.name),
      type: 'scatter', mode: 'markers+text',
      marker: {{
        size: bubbleData.map(d => Math.max(12, Math.min(80, d.roi * 5))),
        color: bubbleData.map(d => d.cac),
        colorscale: 'RdYlGn_r',
        showscale: true,
        colorbar: {{title: 'CAC(¥)', len:0.6, thickness:10}},
        line: {{color:'white', width:1}}
      }},
      textposition: 'top center',
      textfont: {{size:10}},
      hovertemplate: '%{{text}}<br>消耗: ¥%{{x:,.2f}}<br>新增会员: %{{y:,}}<br>ROI: %{{marker.size/5:.1f}}<br>CAC: ¥%{{marker.color:.1f}}<extra></extra>'
    }], {{
      margin:{{l:50,r:50,t:10,b:40}},
      xaxis:{{title:'消耗(¥)', type:'log', showgrid:true, gridcolor:'#f0f0f0'}},
      yaxis:{{title:'新增会员', showgrid:true, gridcolor:'#f0f0f0', rangemode:'tozero'}},
      hovermode:'closest',
      paper_bgcolor:'white', plot_bgcolor:'white'
    }}, {{responsive:true,displayModeBar:false}});
    
    const bestName = bubbleData.reduce((a,b) => a.roi > b.roi ? a : b).name;
    const worstName = bubbleData.reduce((a,b) => a.roi < b.roi ? a : b).name;
    document.getElementById('insight5').textContent =
      `💡 气泡大小=ROI（越大越高效）| 颜色越绿CAC越低 | ROI最高: ${bestName} | ROI最低: ${worstName}`;
  }}
  
  // ---- Chart 3: GMV 贡献排行 ----
  const gmvNames = [...dd.names].reverse();
  const gmvVals = [...dd.gmv].reverse();
  const gmvTop15 = gmvNames.slice(-15);
  const gmvVals15 = gmvVals.slice(-15);
  
  Plotly.newPlot('c3', [
    {{x: gmvVals15, y: gmvTop15, type:'bar', orientation:'h',
      marker:{{color: gmvVals15.map(v => v > gmvVals15[Math.floor(gmvVals15.length/2)] ? C.green : '#86efac'), opacity:0.85}},
      hovertemplate:'%{{y}}<br>GMV: ¥%{{x:,.2f}}<extra></extra>'}}
  ], {{
    margin:{{l:70,r:20,t:5,b:40}}, hovermode:'y unified',
    xaxis:{{title:'GMV(¥)', showgrid:true, gridcolor:'#f0f0f0'}},
    yaxis:{{tickfont:{{size:11}}, automargin:true}},
    paper_bgcolor:'white', plot_bgcolor:'white'
  }}, {{responsive:true,displayModeBar:false}});
  
  document.getElementById('insight3').textContent =
    `💡 TOP1 GMV: ${dd.names[0]} (¥${(dd.gmv[0]/10000).toFixed(1)}万) | TOP3合计: ¥${((dd.gmv[0]+dd.gmv[1]+dd.gmv[2])/10000).toFixed(1)}万`;
}}

// Init
applyFilter();
</script>
</body>
</html>'''

with open('/Users/edy/Desktop/brand_member_dashboard_v2.html', 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Written: {len(html)} bytes to brand_member_dashboard_v2.html")
