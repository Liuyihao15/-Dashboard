#!/usr/bin/env python3
"""
1. Restore old JSON (49fc8b2 - 53 days up to 6/22)
2. Parse new xlsx (6/23-6/28) with CORRECT column mapping (N=13 for trade_members)
3. Generate correct merged JSON
"""
import json, zipfile, xml.etree.ElementTree as ET, os, subprocess

DASHBOARD_DIR = '/Users/edy/Desktop/霸王茶姬看板'
DASHBOARD_DATA = os.path.join(DASHBOARD_DIR, 'dashboard_data.json')
BUILD_SCRIPT = os.path.join(DASHBOARD_DIR, '_build.sh')
XLSX_PATH = "/Users/edy/.hermes/cache/documents/doc_346728300b31_广告会员数据&全量注册会员6月23日-28日.xlsx"

# === Parse xlsx ===
z = zipfile.ZipFile(XLSX_PATH)
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
xl_shared = z.read('xl/sharedStrings.xml')
st = ET.fromstring(xl_shared)
strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' 
           for si in st.findall('.//s:si', ns)]

sheet2 = z.read('xl/worksheets/sheet2.xml')
root = ET.fromstring(sheet2)
rows = root.findall('.//s:row', ns)

def safe_float(v):
    if v is None or str(v).strip().upper() in ('NULL','NONE','','NAN','INF','-INF'):
        return 0.0
    try: return float(v)
    except: return 0.0

# Column mapping (0-based): J=9(K=10,N=13,W=22,X=23)
COL_COST = 9; COL_MEMBERS = 10; COL_TRADE = 13
COL_ORDER_USERS = 22; COL_ORG_GMV = 23
COL_TARGET = 1  # B

# Sum by date with NEW and RET separated
date_data = {}
for r in rows[1:]:
    vals = {}
    for c in r.findall('s:c', ns):
        col_str = c.get('r')
        letter = col_str[0]
        typ = c.get('t', '')
        v = c.find('s:v', ns)
        if v is not None:
            if typ == 's':
                idx = int(v.text)
                vals[letter] = strings[idx] if idx < len(strings) else f"[{idx}]"
            else:
                vals[letter] = v.text
    
    date_str = str(vals.get('C', '')).strip()
    if not date_str:
        continue
    
    target_type = str(vals.get('B', ''))
    is_acq = '新会员' in target_type
    
    cost = safe_float(vals.get('J'))
    members = safe_float(vals.get('K'))
    trade = safe_float(vals.get('N'))
    order_users = safe_float(vals.get('W'))
    org_gmv = safe_float(vals.get('X'))
    
    if date_str not in date_data:
        date_data[date_str] = {
            'acq_cost':0,'ret_cost':0,
            'acq_members':0,'ret_members':0,
            'acq_trade':0,'ret_trade':0,
            'acq_order_users':0,'ret_order_users':0,
            'org_gmv':0
        }
    
    date_data[date_str]['org_gmv'] += org_gmv
    if is_acq:
        date_data[date_str]['acq_cost'] += cost
        date_data[date_str]['acq_members'] += members
        date_data[date_str]['acq_trade'] += trade
        date_data[date_str]['acq_order_users'] += order_users
    else:
        date_data[date_str]['ret_cost'] += cost
        date_data[date_str]['ret_members'] += members
        date_data[date_str]['ret_trade'] += trade
        date_data[date_str]['ret_order_users'] += order_users

# Province aggregation
prov_stats = {}
city_stats = {}
CITY_TO_PROVINCE = {
    '北京':'北京','上海':'上海','天津':'天津','重庆':'重庆',
    '广州':'广东','深圳':'广东','东莞':'广东','佛山':'广东','珠海':'广东','中山':'广东','惠州':'广东',
    '汕头':'广东','江门':'广东','湛江':'广东','肇庆':'广东','茂名':'广东','梅州':'广东','揭阳':'广东',
    '韶关':'广东','清远':'广东','阳江':'广东','潮州':'广东','河源':'广东','云浮':'广东','汕尾':'广东','普宁':'广东',
    '南京':'江苏','苏州':'江苏','无锡':'江苏','常州':'江苏','徐州':'江苏','南通':'江苏','扬州':'江苏',
    '盐城':'江苏','镇江':'江苏','泰州':'江苏','淮安':'江苏','连云港':'江苏','宿迁':'江苏','昆山':'江苏',
    '张家港':'江苏','常熟':'江苏','太仓':'江苏','宜兴':'江苏','江阴':'江苏','丹阳':'江苏','靖江':'江苏',
    '杭州':'浙江','宁波':'浙江','温州':'浙江','金华':'浙江','嘉兴':'浙江','绍兴':'浙江','台州':'浙江',
    '湖州':'浙江','衢州':'浙江','丽水':'浙江','舟山':'浙江','义乌':'浙江','乐清':'浙江','瑞安':'浙江',
    '慈溪':'浙江','诸暨':'浙江','上虞':'浙江','海宁':'浙江','桐乡':'浙江','东阳':'浙江','永康':'浙江',
    '成都':'四川','绵阳':'四川','德阳':'四川','宜宾':'四川','南充':'四川','泸州':'四川','达州':'四川',
    '乐山':'四川','眉山':'四川','自贡':'四川','资阳':'四川','广元':'四川','内江':'四川','遂宁':'四川',
    '武汉':'湖北','宜昌':'湖北','襄阳':'湖北','荆州':'湖北','黄冈':'湖北','孝感':'湖北','十堰':'湖北',
    '黄石':'湖北','咸宁':'湖北','荆门':'湖北','恩施':'湖北',
    '长沙':'湖南','衡阳':'湖南','株洲':'湖南','湘潭':'湖南','岳阳':'湖南','常德':'湖南','郴州':'湖南',
    '益阳':'湖南','邵阳':'湖南','永州':'湖南','怀化':'湖南','娄底':'湖南','湘西':'湖南','张家界':'湖南',
    '福州':'福建','厦门':'福建','泉州':'福建','漳州':'福建','莆田':'福建','龙岩':'福建','宁德':'福建',
    '三明':'福建','南平':'福建',
    '郑州':'河南','洛阳':'河南','南阳':'河南','新乡':'河南','安阳':'河南','开封':'河南','许昌':'河南',
    '平顶山':'河南','周口':'河南','商丘':'河南','焦作':'河南','信阳':'河南','濮阳':'河南','漯河':'河南',
    '驻马店':'河南','三门峡':'河南','鹤壁':'河南',
    '济南':'山东','青岛':'山东','烟台':'山东','潍坊':'山东','临沂':'山东','淄博':'山东','济宁':'山东',
    '威海':'山东','泰安':'山东','日照':'山东','东营':'山东','德州':'山东','滨州':'山东','枣庄':'山东',
    '菏泽':'山东','聊城':'山东','莱芜':'山东',
    '西安':'陕西','咸阳':'陕西','宝鸡':'陕西','榆林':'陕西','汉中':'陕西','渭南':'陕西','安康':'陕西',
    '石家庄':'河北','唐山':'河北','保定':'河北','邯郸':'河北','廊坊':'河北','沧州':'河北','秦皇岛':'河北',
    '邢台':'河北','衡水':'河北','张家口':'河北','承德':'河北',
    '昆明':'云南','曲靖':'云南','大理':'云南','红河':'云南','玉溪':'云南','楚雄':'云南','文山':'云南',
    '合肥':'安徽','芜湖':'安徽','马鞍山':'安徽','安庆':'安徽','滁州':'安徽','蚌埠':'安徽','阜阳':'安徽',
    '六安':'安徽','宿州':'安徽','黄山':'安徽','宣城':'安徽','淮南':'安徽','铜陵':'安徽','池州':'安徽',
    '南宁':'广西','柳州':'广西','桂林':'广西','玉林':'广西','北海':'广西','梧州':'广西','贵港':'广西',
    '百色':'广西','钦州':'广西','防城港':'广西','河池':'广西','来宾':'广西','贺州':'广西','崇左':'广西',
    '南昌':'江西','赣州':'江西','九江':'江西','宜春':'江西','上饶':'江西','吉安':'江西','抚州':'江西',
    '景德镇':'江西','萍乡':'江西','新余':'江西','鹰潭':'江西',
    '太原':'山西','大同':'山西','运城':'山西','临汾':'山西','晋中':'山西','长治':'山西','吕梁':'山西',
    '晋城':'山西','阳泉':'山西','忻州':'山西','朔州':'山西',
    '哈尔滨':'黑龙江','大庆':'黑龙江','齐齐哈尔':'黑龙江','牡丹江':'黑龙江','佳木斯':'黑龙江',
    '长春':'吉林','吉林':'吉林','延边':'吉林','四平':'吉林','松原':'吉林','通化':'吉林',
    '沈阳':'辽宁','大连':'辽宁','鞍山':'辽宁','锦州':'辽宁','抚顺':'辽宁','营口':'辽宁','丹东':'辽宁',
    '盘锦':'辽宁','朝阳':'辽宁','辽阳':'辽宁','本溪':'辽宁','阜新':'辽宁','葫芦岛':'辽宁','铁岭':'辽宁',
    '兰州':'甘肃','天水':'甘肃','庆阳':'甘肃','酒泉':'甘肃','武威':'甘肃','定西':'甘肃',
    '贵阳':'贵州','遵义':'贵州','黔东南':'贵州','黔南':'贵州','六盘水':'贵州',
    '呼和浩特':'内蒙古','包头':'内蒙古','鄂尔多斯':'内蒙古','赤峰':'内蒙古','通辽':'内蒙古',
    '海口':'海南','三亚':'海南','儋州':'海南',
    '银川':'宁夏','吴忠':'宁夏','中卫':'宁夏','石嘴山':'宁夏',
    '西宁':'青海','海东':'青海',
    '拉萨':'西藏','日喀则':'西藏',
    '乌鲁木齐':'新疆','昌吉':'新疆','伊犁':'新疆','克拉玛依':'新疆','阿克苏':'新疆','喀什':'新疆',
}

def resolve_province(city_name):
    if not city_name: return None
    city_name = city_name.strip()
    if city_name in CITY_TO_PROVINCE: return CITY_TO_PROVINCE[city_name]
    provinces = ['广东','江苏','浙江','广西','福建','江西','湖南','湖北','河南','河北','山东',
                 '山西','陕西','四川','贵州','云南','安徽','甘肃','青海','西藏','海南',
                 '辽宁','吉林','黑龙江','内蒙古','新疆','宁夏','北京','上海','天津','重庆']
    for p in provinces:
        if p in city_name: return p
    for c, p in sorted(CITY_TO_PROVINCE.items(), key=lambda x: -len(x[0])):
        if c in city_name or city_name in c: return p
    return None

for r in rows[1:]:
    vals = {}
    for c in r.findall('s:c', ns):
        col_str = c.get('r'); letter = col_str[0]; typ = c.get('t', '')
        v = c.find('s:v', ns)
        if v is not None:
            if typ == 's': vals[letter] = strings[int(v.text)]
            else: vals[letter] = v.text
    
    cost = safe_float(vals.get('J'))
    members = safe_float(vals.get('K'))
    org_gmv = safe_float(vals.get('X'))
    trade = safe_float(vals.get('N'))
    order_users = safe_float(vals.get('W'))
    province = vals.get('F', '').strip()
    city = vals.get('G', '').strip()
    
    if not province or province in ('0',''): province = resolve_province(city) or ''
    if not city: continue
    
    if province:
        if province not in prov_stats:
            prov_stats[province] = {'cost':0,'members':0,'org_gmv':0,'order_users':0,'trade_members':0}
        prov_stats[province]['cost'] += cost
        prov_stats[province]['members'] += members
        prov_stats[province]['org_gmv'] += org_gmv
        prov_stats[province]['order_users'] += order_users
        prov_stats[province]['trade_members'] += trade
    
    if city:
        if city not in city_stats:
            city_stats[city] = {'cost':0,'members':0,'org_gmv':0}
        city_stats[city]['cost'] += cost
        city_stats[city]['members'] += members
        city_stats[city]['org_gmv'] += org_gmv

print("=== 原始xlsx 6/23-6/28 ===")
total_trade = 0
new_dates = sorted(date_data.keys())
for date in new_dates:
    dd = date_data[date]
    t = dd['acq_trade'] + dd['ret_trade']
    total_trade += t
    print(f"{date}: cost={dd['acq_cost']+dd['ret_cost']:.0f} members={dd['acq_members']:.0f} trade={t:.0f}(acq={dd['acq_trade']:.0f}+ret={dd['ret_trade']:.0f}) org={dd['org_gmv']:.0f}")
print(f"合计 trade_members: {total_trade:.0f}")

# === Load existing JSON ===
with open(DASHBOARD_DATA) as f: existing = json.load(f)
print(f"\n现有: {existing['dates'][0]}~{existing['dates'][-1]} ({len(existing['dates'])}天)")

# === Merge ===
merged = dict(existing)

# Daily fields to extend
fields_map = {
    'acq_cost': ('acq_cost', 0),
    'acq_members': ('acq_members', 0),
    'acq_trade_members': ('acq_trade', 0),
    'acq_order_users': ('acq_order_users', 0),
    'ret_cost': ('ret_cost', 0),
    'ret_members': ('ret_members', 0),
    'ret_trade_members': ('ret_trade', 0),
    'ret_order_users': ('ret_order_users', 0),
    'org_gmv': ('org_gmv', 0),
}

for json_field, (src_field, default) in fields_map.items():
    if json_field not in merged:
        merged[json_field] = [default] * len(merged['dates'])
    for d in new_dates:
        merged[json_field].append(int(date_data[d][src_field]) if src_field in ('acq_members','ret_members','acq_trade','ret_trade','acq_order_users','ret_order_users') else round(date_data[d][src_field], 2) if src_field in ('acq_cost','ret_cost','org_gmv') else date_data[d].get(src_field, default))

merged['dates'] = merged['dates'] + new_dates

# Province merge
all_provs = list(merged.get('provinces', []))
for p in prov_stats:
    if p not in all_provs: all_provs.append(p)

prov_fields = ['prov_cost','prov_members','prov_org_gmv','prov_order_users','prov_trade_members']
prov_data = {}
for i, p in enumerate(merged.get('provinces', [])):
    prov_data[p] = {f: merged.get(f, [0]*len(all_provs))[i] for f in prov_fields}
for p, stats in prov_stats.items():
    if p not in prov_data:
        prov_data[p] = {f: 0 for f in prov_fields}
    prov_data[p]['prov_cost'] += stats['cost']
    prov_data[p]['prov_members'] += stats['members']
    prov_data[p]['prov_org_gmv'] += stats['org_gmv']
    prov_data[p]['prov_order_users'] += stats['order_users']
    prov_data[p]['prov_trade_members'] += stats['trade_members']

prov_sorted = sorted(prov_data.keys(), key=lambda p: prov_data[p]['prov_cost'], reverse=True)
merged['provinces'] = prov_sorted
for f in prov_fields:
    merged[f] = [prov_data[p][f] for p in prov_sorted]
merged['prov_cac'] = [round(merged['prov_cost'][i]/merged['prov_members'][i], 2) if merged['prov_members'][i]>0 else 0 for i in range(len(prov_sorted))]
merged['prov_roi'] = [round(merged['prov_org_gmv'][i]/merged['prov_cost'][i], 2) if merged['prov_cost'][i]>0 else 0 for i in range(len(prov_sorted))]

# City merge
all_cities = list(set(merged.get('cities', []) + list(city_stats.keys())))
city_cost = {c: merged.get('city_cost',[0])[merged['cities'].index(c)] if c in merged.get('cities',[]) else 0 for c in all_cities}
city_members = {c: merged.get('city_members',[0])[merged['cities'].index(c)] if c in merged.get('cities',[]) else 0 for c in all_cities}
city_gmv = {c: merged.get('city_gmv',[0])[merged['cities'].index(c)] if c in merged.get('cities',[]) else 0 for c in all_cities}
for c, stats in city_stats.items():
    city_cost[c] += stats['cost']
    city_members[c] += stats['members']
    city_gmv[c] += stats['org_gmv']

city_sorted = sorted(all_cities, key=lambda c: city_cost[c], reverse=True)
merged['cities'] = city_sorted
merged['city_cost'] = [round(city_cost[c], 2) for c in city_sorted]
merged['city_members'] = [int(city_members[c]) for c in city_sorted]
merged['city_gmv'] = [round(city_gmv[c], 2) for c in city_sorted]
merged['city_roi'] = [round(city_gmv[c]/city_cost[c], 2) if city_cost[c]>0 else 0 for c in city_sorted]
merged['city_cac'] = [round(city_cost[c]/city_members[c], 2) if city_members[c]>0 else 0 for c in city_sorted]

existing_city_prov = dict(zip(merged.get('cities',[]), merged.get('city_prov',[])))
for c in city_sorted:
    if c not in existing_city_prov: existing_city_prov[c] = resolve_province(c) or ''
merged['city_prov'] = [existing_city_prov.get(c, '') for c in city_sorted]

# Save
with open(DASHBOARD_DATA, 'w', encoding='utf-8') as f:
    json.dump(merged, f, ensure_ascii=False, separators=(',',':'))

# Verify
print(f"\n=== 验证修正后数据 ===")
for date in ['20260623','20260624','20260625','20260626','20260627','20260628']:
    idx = merged['dates'].index(date)
    at = merged['acq_trade_members'][idx]
    rt = merged['ret_trade_members'][idx]
    ts = at + rt
    c = merged['acq_cost'][idx] + merged['ret_cost'][idx]
    og = merged['org_gmv'][idx]
    print(f"{date}: cost={c:.0f} trade_members={ts}(acq={at}+ret={rt}) org={og:.0f} roi={og/c:.2f}")

total_t = sum(merged.get('acq_trade_members',[])) + sum(merged.get('ret_trade_members',[]))
total_c = sum(merged.get('acq_cost',[])) + sum(merged.get('ret_cost',[]))
total_o = sum(merged.get('org_gmv',[]))
print(f"\n总量: cost={total_c:.0f} trade_members={total_t} org={total_o:.0f} roi={total_o/total_c:.2f}")
print(f"\n✅ JSON saved")
