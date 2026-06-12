#!/usr/bin/env python3
"""Generate brand member dashboard v2"""
import zipfile, re, json, datetime
from xml.etree import ElementTree as ET
from collections import defaultdict, Counter

def col_letter_to_index(col_str):
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

print("Reading data...")
with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    daily = {}
    for idx, row_xml in enumerate(rows[1:]):
        if idx % 50000 == 0: print(f"  Row {idx}...")
        cells = re.findall(r'(<c[^>]*>.*?</c>)', row_xml, re.DOTALL)
        row = {}
        for cell_str in cells:
            rm = re.search(r'r="([^"]+)"', cell_str)
            if not rm: continue
            ref = rm.group(1)
            m2 = re.match(r'([A-Z]+)', ref)
            if not m2: continue
            col = 0
            for c in m2.group(1):
                col = col * 26 + (ord(c) - ord('A') + 1)
            col -= 1
            is_str = 't="s"' in cell_str
            vm = re.search(r'<v>(.*?)</v>', cell_str)
            if vm:
                val = vm.group(1)
                if is_str:
                    try: val = strings[int(val)]
                    except: pass
                row[col] = val
        
        date = row.get(2)
        if not date: continue
        if date not in daily:
            daily[date] = {'cost':0.0,'new_members':0,'gmv':0.0,'orders':0,'clicks':0,'impr':0,
                          'type_cost':defaultdict(float),'type_members':defaultdict(int),
                          'city_cost':defaultdict(float),'city_members':defaultdict(int),'city_gmv':defaultdict(float)}
        d = daily[date]
        cost = float(row.get(9) or 0)
        d['cost'] += cost
        k_val = row.get(10)
        if k_val and k_val not in ('NULL','None',''):
            try: d['new_members'] += int(float(k_val))
            except: pass
        d['gmv'] += float(row.get(24) or 0)
        d['orders'] += int(float(row.get(18) or 0))
        d['clicks'] += int(float(row.get(14) or 0))
        d['impr'] += int(float(row.get(15) or 0))
        t = row.get(1)
        if t:
            d['type_cost'][t] += cost
            if k_val and k_val not in ('NULL','None',''):
                try: d['type_members'][t] += int(float(k_val))
                except: pass
        city = row.get(5)
        if city and city not in ('0','NULL','',None):
            d['city_cost'][city] += cost
            if k_val and k_val not in ('NULL','None',''):
                try: d['city_members'][city] += int(float(k_val))
                except: pass
            d['city_gmv'][city] += float(row.get(24) or 0)

sorted_dates = sorted(daily.keys())
dates_arr = []; cost_arr = []; member_arr = []; gmv_arr = []; order_arr = []; roi_arr = []; cpa_arr = []
for d in sorted_dates:
    dd = daily[d]
    dates_arr.append(d); cost_arr.append(round(dd['cost'],2)); member_arr.append(dd['new_members'])
    gmv_arr.append(round(dd['gmv'],2)); order_arr.append(dd['orders'])
    roi_arr.append(round(dd['gmv']/dd['cost'],2) if dd['cost']>0 else 0)
    cpa_arr.append(round(dd['cost']/dd['new_members'],2) if dd['new_members']>0 else 0)

print(f"Members: {sum(member_arr)} | Cost: {sum(cost_arr):.2f} | GMV: {sum(gmv_arr):.2f} | Orders: {sum(order_arr)}")

# Weekly
week_map = defaultdict(lambda:{'cost':0,'members':0,'gmv':0,'orders':0})
for ds,dd in daily.items():
    dt = datetime.datetime.strptime(ds,'%Y%m%d'); wk = dt.strftime('%Y-W%V')
    week_map[wk]['cost']+=dd['cost']; week_map[wk]['members']+=dd['new_members']
    week_map[wk]['gmv']+=dd['gmv']; week_map[wk]['orders']+=dd['orders']
wk_l = sorted(week_map.keys()); wk_c = [round(week_map[w]['cost'],2) for w in wk_l]
wk_m = [week_map[w]['members'] for w in wk_l]; wk_g = [round(week_map[w]['gmv'],2) for w in wk_l]

# Monthly
month_map = defaultdict(lambda:{'cost':0,'members':0,'gmv':0,'orders':0})
for ds,dd in daily.items():
    mo = ds[:6]; month_map[mo]['cost']+=dd['cost']; month_map[mo]['members']+=dd['new_members']
    month_map[mo]['gmv']+=dd['gmv']; month_map[mo]['orders']+=dd['orders']
mo_l = sorted(month_map.keys()); mo_c = [round(month_map[m]['cost'],2) for m in mo_l]
mo_mem = [month_map[m]['members'] for m in mo_l]; mo_g = [round(month_map[m]['gmv'],2) for m in mo_l]

# City top20
city_cost_total=Counter(); city_member_total=Counter(); city_gmv_total=Counter()
for dd in daily.values():
    for c,v in dd['city_cost'].items(): city_cost_total[c]+=v
    for c,v in dd['city_members'].items(): city_member_total[c]+=v
    for c,v in dd['city_gmv'].items(): city_gmv_total[c]+=v
top_c = [c for c,_ in city_cost_total.most_common(20)]
top_cst = [round(city_cost_total[c],2) for c in top_c]
top_mem = [city_member_total[c] for c in top_c]
top_roi = [round(city_gmv_total[c]/city_cost_total[c],2) if city_cost_total[c]>0 else 0 for c in top_c]

# Target type
acq_cost = round(sum(dd['type_cost'].get('27',0) for dd in daily.values()),2)
acq_members = sum(dd['type_members'].get('27',0) for dd in daily.values())
ret_cost = round(sum(dd['type_cost'].get('32',0) for dd in daily.values()),2)
ret_members = sum(dd['type_members'].get('32',0) for dd in daily.values())

D = {
    'dates':dates_arr,'cost':cost_arr,'members':member_arr,'gmv':gmv_arr,'orders':order_arr,
    'roi':roi_arr,'cpa':cpa_arr,
    'wkLabels':wk_l,'wkCost':wk_c,'wkMembers':wk_m,'wkGmv':wk_g,
    'moLabels':mo_l,'moCost':mo_c,'moMembers':mo_mem,'moGmv':mo_g,
    'cityNames':top_c,'cityCosts':top_cst,'cityMembers':top_mem,'cityRois':top_roi,
    'acqCost':acq_cost,'acqMembers':acq_members,'retCost':ret_cost,'retMembers':ret_members
}
data_json = json.dumps(D, ensure_ascii=False)

period = '%s~%s (%d天)' % (sorted_dates[0][:4]+'-'+sorted_dates[0][4:6]+'-'+sorted_dates[0][6:8],
                            sorted_dates[-1][:4]+'-'+sorted_dates[-1][4:6]+'-'+sorted_dates[-1][6:8],
                            len(sorted_dates))

html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>霸王茶姬 · 品牌会员投放 Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,"PingFang SC","Helvetica Neue",sans-serif;background:#f5f7fa;color:#1a1a2e;padding:20px;max-width:1400px;margin:0 auto}
.header{display:flex;justify-content:space-between;align-items:center;margin-bottom:24px;flex-wrap:wrap;gap:12px}
.header h1{font-size:22px;font-weight:700;color:#1a1a2e}
.header .sub{color:#6b7280;font-size:14px}
.filters{display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap;align-items:flex-end}
.filters .fg{display:flex;flex-direction:column}
.filters label{font-size:12px;font-weight:600;color:#374151;margin-bottom:2px}
.filters select,.filters input{padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;background:#fff;cursor:pointer}
.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:20px}
.kpi-card{background:#fff;border-radius:10px;padding:14px 16px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
.kpi-card .l{font-size:11px;color:#6b7280;margin-bottom:2px}
.kpi-card .v{font-size:20px;font-weight:700;color:#1a1a2e}
.chart-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:20px}
.chart-full{grid-column:1/-1}
.chart-box{background:#fff;border-radius:10px;padding:10px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
.chart-box h3{font-size:13px;font-weight:600;color:#374151;margin-bottom:6px;padding-left:2px}
@media(max-width:900px){.chart-grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="header">
  <div><h1>\U0001F956 霸王茶姬 · 品牌会员投放 Dashboard</h1><div class="sub">""" + period + """</div></div>
</div>
<div class="filters">
  <div class="fg"><label>时间粒度</label><select id="grain" onchange="redraw()"><option value="day">按日</option><option value="week">按周</option><option value="month">按月</option></select></div>
  <div class="fg"><label>开始</label><input type="date" id="ds" onchange="redraw()"></div>
  <div class="fg"><label>结束</label><input type="date" id="de" onchange="redraw()"></div>
  <div class="fg"><label>目标类型</label><select id="target" onchange="redraw()"><option value="all">全部</option><option value="acq">新会员注册</option><option value="ret">老会员转化</option></select></div>
  <div class="fg"><label>城市</label><select id="city" onchange="redraw()"><option value="all">全部城市</option></select></div>
</div>
<div class="kpi-grid" id="kpiGrid"></div>
<div class="chart-grid">
  <div class="chart-full chart-box"><h3>\U0001F4C8 消耗 & 新增会员趋势</h3><div id="c1" style="height:340px;"></div></div>
  <div class="chart-box"><h3>\U0001F4CA CPA & ROI 趋势</h3><div id="c2" style="height:300px;"></div></div>
  <div class="chart-box"><h3>\U0001F504 新增会员 vs 订单数</h3><div id="c3" style="height:300px;"></div></div>
  <div class="chart-box"><h3>\U0001F3C6 TOP20 城市消耗排行</h3><div id="c4" style="height:340px;"></div></div>
  <div class="chart-box"><h3>\U0001F4CC 城市效率矩阵</h3><div id="c5" style="height:340px;"></div></div>
  <div class="chart-box"><h3>\U0001F4CA 拉新 vs 复购</h3><div id="c6" style="height:300px;"></div></div>
</div>
<script>
const D = """ + data_json + """;
D.cityNames.forEach(function(c){var o=document.createElement('option');o.value=c;o.textContent=c;document.getElementById('city').appendChild(o)});
var fd=D.dates[0],ld=D.dates[D.dates.length-1];
document.getElementById('ds').value=fd.slice(0,4)+'-'+fd.slice(4,6)+'-'+fd.slice(6,8);
document.getElementById('de').value=ld.slice(0,4)+'-'+ld.slice(4,6)+'-'+ld.slice(6,8);
var C={prim:'#2563eb',green:'#10b981',red:'#ef4444',yellow:'#f59e0b',purple:'#8b5cf6'};
function fi(){var s=document.getElementById('ds').value.replace(/-/g,''),e=document.getElementById('de').value.replace(/-/g,'');
  var si=0,ei=D.dates.length-1;
  for(var i=0;i<D.dates.length;i++){if(D.dates[i]>=s){si=i;break}}
  for(var i=D.dates.length-1;i>=0;i--){if(D.dates[i]<=e){ei=i;break}}
  return[si,ei]}
function uk(){var s=fi(),c=D.cost.slice(s[0],s[1]+1).reduce(function(a,b){return a+b},0);
  var m=D.members.slice(s[0],s[1]+1).reduce(function(a,b){return a+b},0);
  var g=D.gmv.slice(s[0],s[1]+1).reduce(function(a,b){return a+b},0);
  var o=D.orders.slice(s[0],s[1]+1).reduce(function(a,b){return a+b},0);
  document.getElementById('kpiGrid').innerHTML='<div class=kpi-card><div class=l>总消耗</div><div class=v>'+c.toFixed(2)+'</div></div><div class=kpi-card><div class=l>新增会员</div><div class=v>'+m.toLocaleString()+'</div></div><div class=kpi-card><div class=l>CPA</div><div class=v>'+(m?(c/m).toFixed(2):'0.00')+'</div></div><div class=kpi-card><div class=l>广告 ROI</div><div class=v>'+(c?(g/c).toFixed(2):'0.00')+'</div></div><div class=kpi-card><div class=l>实付 GMV</div><div class=v>'+g.toFixed(2)+'</div></div><div class=kpi-card><div class=l>订单数</div><div class=v>'+o.toLocaleString()+'</div></div>'}
function redraw(){uk();var g=document.getElementById('grain').value,s=fi();
  var x,cost,members,gmv,orders,roi,cpa;
  if(g==='week'&&D.wkLabels.length){x=D.wkLabels;cost=D.wkCost;members=D.wkMembers;gmv=D.wkGmv;
    roi=cost.map(function(c,i){return gmv[i]/c||0});cpa=cost.map(function(c,i){return members[i]?c/members[i]:0})}
  else if(g==='month'&&D.moLabels.length){x=D.moLabels;cost=D.moCost;members=D.moMembers;gmv=D.moGmv;
    roi=cost.map(function(c,i){return gmv[i]/c||0});cpa=cost.map(function(c,i){return members[i]?c/members[i]:0})}
  else{x=D.dates.slice(s[0],s[1]+1);cost=D.cost.slice(s[0],s[1]+1);members=D.members.slice(s[0],s[1]+1);
    gmv=D.gmv.slice(s[0],s[1]+1);orders=D.orders.slice(s[0],s[1]+1);roi=D.roi.slice(s[0],s[1]+1);cpa=D.cpa.slice(s[0],s[1]+1)}
  Plotly.newPlot('c1',[{x,y:cost,type:'scatter',mode:'lines+markers',name:'消耗',line:{color:C.prim,width:2},marker:{size:4},yaxis:'y',hovertemplate:'%{x}<br>消耗: \\u00a5%{y:,.2f}<extra></extra>'},{x,y:members,type:'scatter',mode:'lines+markers',name:'新增会员',line:{color:C.green,width:2},marker:{size:4},yaxis:'y2',hovertemplate:'%{x}<br>新增: %{y:,}\\u4eba<extra></extra>'}],{margin:{l:60,r:60,t:10,b:40},hovermode:'x unified',xaxis:{tickangle:-45,tickfont:{size:10}},yaxis:{title:'\\u6d88\\u8017 (\\u00a5)',side:'left',showgrid:true,gridcolor:'#f0f0f0'},yaxis2:{title:'\\u65b0\\u589e\\u4f1a\\u5458',side:'right',overlaying:'y',showgrid:false},legend:{orientation:'h',y:1.08},paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})
  Plotly.newPlot('c2',[{x,y:cpa,type:'scatter',mode:'lines+markers',name:'CPA',line:{color:C.red,width:2},marker:{size:4},yaxis:'y',hovertemplate:'%{x}<br>CPA: \\u00a5%{y:,.2f}<extra></extra>'},{x,y:roi,type:'scatter',mode:'lines+markers',name:'ROI',line:{color:C.yellow,width:2},marker:{size:4},yaxis:'y2',hovertemplate:'%{x}<br>ROI: %{y:.2f}<extra></extra>'}],{margin:{l:50,r:50,t:10,b:40},hovermode:'x unified',xaxis:{tickangle:-45,tickfont:{size:10}},yaxis:{title:'CPA (\\u00a5)',side:'left',showgrid:true,gridcolor:'#f0f0f0'},yaxis2:{title:'ROI',side:'right',overlaying:'y',showgrid:false,rangemode:'tozero'},legend:{orientation:'h',y:1.08},paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})
  Plotly.newPlot('c3',[{x,y:members,type:'bar',name:'新增会员',marker:{color:C.green,opacity:0.7},yaxis:'y',hovertemplate:'%{x}<br>新增: %{y:,}\\u4eba<extra></extra>'},{x,y:orders,type:'scatter',mode:'lines+markers',name:'订单数',line:{color:C.purple,width:2},marker:{size:4},yaxis:'y2',hovertemplate:'%{x}<br>订单: %{y:,}<extra></extra>'}],{margin:{l:50,r:50,t:10,b:40},hovermode:'x unified',xaxis:{tickangle:-45,tickfont:{size:10}},yaxis:{title:'\\u65b0\\u589e\\u4f1a\\u5458',side:'left',showgrid:true,gridcolor:'#f0f0f0'},yaxis2:{title:'\\u8ba2\\u5355\\u6570',side:'right',overlaying:'y',showgrid:false},legend:{orientation:'h',y:1.08},paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})
  Plotly.newPlot('c4',[{x:D.cityNames,y:D.cityCosts,type:'bar',marker:{color:D.cityCosts.map(function(c){return c>D.cityCosts[4]?C.prim:'#93c5fd'}),opacity:0.85},hovertemplate:'%{x}<br>\\u6d88\\u8017: \\u00a5%{y:,.2f}<extra></extra>'}],{margin:{l:50,r:20,t:10,b:80},hovermode:'x unified',xaxis:{tickangle:-45,tickfont:{size:10}},yaxis:{title:'\\u6d88\\u8017 (\\u00a5)',showgrid:true,gridcolor:'#f0f0f0'},paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})
  Plotly.newPlot('c5',[{x:D.cityCosts,y:D.cityRois,text:D.cityNames,type:'scatter',mode:'markers+text',marker:{size:D.cityMembers.map(function(m){return Math.max(10,Math.min(60,m/500))}),color:D.cityRois,colorscale:'RdYlGn',showscale:true,colorbar:{title:'ROI',len:0.6,thickness:10}},textposition:'top center',textfont:{size:10},hovertemplate:'%{text}<br>\\u6d88\\u8017: \\u00a5%{x:,.2f}<br>ROI: %{y:.2f}<br>\\u65b0\\u589e: %{marker.size}k<extra></extra>'}],{margin:{l:50,r:50,t:10,b:40},xaxis:{title:'\\u6d88\\u8017 (\\u00a5)',type:'log',showgrid:true,gridcolor:'#f0f0f0'},yaxis:{title:'ROI',showgrid:true,gridcolor:'#f0f0f0',rangemode:'tozero'},shapes:[{type:'line',x0:Math.min(...D.cityCosts),y0:2,x1:Math.max(...D.cityCosts),y1:2,line:{dash:'dash',color:'#d1d5db',width:1}}],paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})
  Plotly.newPlot('c6',[{x:['\\u62c9\\u65b0\\u00b7\\u65b0\\u589e\\u4f1a\\u5458','\\u590d\\u8d2d\\u00b7\\u65b0\\u589e\\u4f1a\\u5458'],y:[D.acqMembers,D.retMembers],type:'bar',name:'\\u65b0\\u589e\\u4f1a\\u5458',marker:{color:[C.green,C.prim]}},{x:['\\u62c9\\u65b0\\u00b7\\u6d88\\u8017','\\u590d\\u8d2d\\u00b7\\u6d88\\u8017'],y:[D.acqCost,D.retCost],type:'bar',name:'\\u6d88\\u8017',marker:{color:[C.prim,C.purple]}}],{barmode:'group',margin:{l:50,r:20,t:10,b:60},hovermode:'x',xaxis:{type:'category'},yaxis:{title:'\\u503c',showgrid:true,gridcolor:'#f0f0f0'},legend:{orientation:'h',y:1.08},paper_bgcolor:'white',plot_bgcolor:'white'},{responsive:true,displayModeBar:false})}
redraw();
</script>
</body>
</html>"""

out_path = '/Users/edy/Desktop/brand_member_dashboard.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

print("Done: %s" % out_path)
print("Members=%d (correct), Cost=%.2f, GMV=%.2f" % (sum(member_arr), sum(cost_arr), sum(gmv_arr)))
