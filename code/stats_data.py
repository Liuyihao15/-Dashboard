#!/usr/bin/env python3
"""读取品牌会员数据 - 完整统计"""
import zipfile, re
from xml.etree import ElementTree as ET
from collections import Counter, defaultdict

def col_letter_to_index(col_str):
    """A->0, B->1, ..., Z->25"""
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

def parse_cell_value(cell_xml, strings):
    """从cell XML中提取值，含共享字符串解析"""
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

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    # 读取 shared strings
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = []
    for si in sis:
        t = si.find('.//s:t', ns)
        strings.append(t.text if t is not None else '')
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    print(f"总数据行数: {len(rows) - 1}")  # 去掉header
    
    # 提取每个cell的列引用和值
    data = []
    for row_xml in rows[1:]:
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        row_vals = {}
        for ref, content in cells:
            col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
            val = parse_cell_value(f'<c>{content}</c>', strings)
            row_vals[col] = val
        
        if row_vals:
            data.append(row_vals)
    
    print(f"解析成功行数: {len(data)}")
    
    # 统计
    dates = set()
    cities = set()
    types = set()
    total_cost = 0
    total_new_members = 0
    total_gmv = 0
    total_orders = 0
    plan_ids = set()
    
    for row in data:
        d = row.get(2)
        if d: dates.add(d)
        
        c = row.get(5)  # 省份
        if c and c not in ('0', 'NULL', '', None): cities.add(c)
        
        t = row.get(1)
        if t and t != '1': types.add(t)
        
        cost = row.get(9)
        if cost:
            try: total_cost += float(cost)
            except: pass
        
        reg = row.get(10)
        if reg:
            try: total_new_members += int(float(reg))
            except: pass
        
        gmv = row.get(24)
        if gmv:
            try: total_gmv += float(gmv)
            except: pass
        
        orders = row.get(18)
        if orders:
            try: total_orders += int(float(orders))
            except: pass
        
        pid = row.get(3)
        if pid: plan_ids.add(pid)
    
    print(f"\n时间范围: {min(dates)} ~ {max(dates)} ({len(dates)}天)")
    print(f"目标类型: {types}")
    print(f"城市数: {len(cities)}")
    print(f"投放计划数: {len(plan_ids)}")
    print(f"总消耗: {total_cost:.2f}")
    print(f"总注册会员: {total_new_members}")
    print(f"总实付GMV: {total_gmv:.2f}")
    print(f"总订单数: {total_orders}")
    
    # 分类型统计
    print("\n=== 按目标类型统计 ===")
    type_stats = defaultdict(lambda: {'cost': 0, 'new_members': 0, 'gmv': 0, 'orders': 0, 'rows': 0})
    for row in data:
        t = row.get(1)
        type_stats[t]['rows'] += 1
        try: type_stats[t]['cost'] += float(row.get(9) or 0)
        except: pass
        try: type_stats[t]['new_members'] += int(float(row.get(10) or 0))
        except: pass
        try: type_stats[t]['gmv'] += float(row.get(24) or 0)
        except: pass
        try: type_stats[t]['orders'] += int(float(row.get(18) or 0))
        except: pass
    
    for t, s in sorted(type_stats.items(), key=lambda x: -x[1]['cost']):
        print(f"  {t}: {s['rows']}行, 消耗{s['cost']:.2f}, 新增会员{s['new_members']}, 订单{s['orders']}, 实付GMV{s['gmv']:.2f}")
    
    # TOP15城市
    print("\n=== TOP15城市（按消耗）===")
    city_cost = Counter()
    for row in data:
        city = row.get(5)
        cost = row.get(9)
        if city and city not in ('0', 'NULL', '', None) and cost:
            try: city_cost[city] += float(cost)
            except: pass
    for city, cost in city_cost.most_common(15):
        print(f"  {city}: {cost:.2f}")
