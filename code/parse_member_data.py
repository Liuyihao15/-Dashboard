#!/usr/bin/env python3
"""Parse brand member data - aggregate by city, show unique cities, target types, and summary stats"""
import zipfile, re
from xml.etree import ElementTree as ET
from collections import defaultdict
import json

path = '/Users/edy/Desktop/品牌会员数据.xlsx'

with zipfile.ZipFile(path, 'r') as z:
    # Read shared strings
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = []
    for si in sis:
        t = si.find('.//s:t', ns)
        strings.append(t.text if t is not None else '')
    
    # Debug: check first few shared strings
    print("=== First 30 shared strings ===")
    for i, s in enumerate(strings[:30]):
        print(f"  [{i:2d}] {repr(s)}")
    
    # Parse XML data
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    print(f"Total rows (excl header): {len(rows)-1}")
    print(f"Total unique strings: {len(strings)}")
    
    # Column mapping (from previous analysis):
    # [0]=业务品牌, [1]=目标类型, [2]=日期, [3]=投放计划ID, [4]=商家ID
    # [5]=城市名(p省份列,实际是城市), [6]=城市编码, [7]=投放计划名, [8]=KA账号
    # [9]=广告消耗(CA), [10]=注册会员数, [11]=会员注册成本(CPA→CA), [12]=ROI
    # [13]=交易会员数, [14]=点击次数, [15]=曝光次数, [16]=注册率
    # [17]=CVR, [18]=订单数, [19]=CPM, [20]=CPC, [21]=CTR
    # [22]=订单人数, [23]=原价GMV, [24]=实付GMV, [25]=计费原价
    
    def get_cell_val(cell_xml):
        """Extract value from a cell XML"""
        v = re.search(r'<v>(.*?)</v>', cell_xml)
        t_attr = re.search(r't="(\w+)"', cell_xml)
        if v:
            val = v.group(1)
            if t_attr and t_attr.group(1) == 's':
                idx = int(val)
                return strings[idx] if idx < len(strings) else val
            return val
        return ''
    
    # Get target type mapping
    target_types = {}
    target_names = {}
    
    # Also collect unique cities
    cities = set()
    
    # Data by target type × city
    # key: (target_type_idx, city_name)
    agg = defaultdict(lambda: {
        'consumption': 0.0,   # 广告消耗 = CA
        'new_members': 0,     # 注册会员数
        'gmv': 0.0,           # 实付GMV
        'roi_sum': 0.0,
        'roi_count': 0,
        'orders': 0,
        'order_people': 0,
        'impressions': 0,
        'clicks': 0,
        'cpc_sum': 0.0,
        'cpc_count': 0,
        'cpm_sum': 0.0,
        'cpm_count': 0,
        'cvr_sum': 0.0,
        'cvr_count': 0,
    })
    
    # Parse actual data rows (skip row 0 which is column names)
    total_rows = len(rows)
    for idx, row_xml in enumerate(rows[1:], 1):
        cells = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        if len(cells) < 26:
            continue
        
        vals = [get_cell_val(c) for c in cells]
        
        # Skip header row (value '0' in col 0)
        if vals[0] in ('0', '1') and vals[1] in ('0', '1'):
            # Check if this is a data row or header
            try:
                test_num = int(vals[0])
                if test_num == 0 and idx > 1:
                    continue
            except:
                pass
        
        try:
            target_type = vals[1]  # "27" or "32"  
            city = vals[5]  # city name
            consumption = float(vals[9]) if vals[9] else 0.0
            new_members = int(float(vals[10])) if vals[10] else 0
            gmv = float(vals[24]) if vals[24] else 0.0
            roi = float(vals[12]) if vals[12] else 0.0
            orders = int(float(vals[18])) if vals[18] else 0
            order_people = int(float(vals[22])) if vals[22] else 0
            impressions = int(float(vals[15])) if vals[15] else 0
            clicks = int(float(vals[14])) if vals[14] else 0
            cpc = float(vals[20]) if vals[20] else 0.0
            cpm = float(vals[19]) if vals[19] else 0.0
            cvr = float(vals[17]) if vals[17] else 0.0
        except (ValueError, IndexError):
            continue
        
        # Build target_type mapping
        target_names[target_type] = target_names.get(target_type, f"Type_{target_type}")
        cities.add(city)
        
        key = (target_type, city)
        agg[key]['consumption'] += consumption
        agg[key]['new_members'] += new_members
        agg[key]['gmv'] += gmv
        if roi > 0:
            agg[key]['roi_sum'] += roi * consumption  # weighted by consumption
            agg[key]['roi_count'] += consumption
        agg[key]['orders'] += orders
        agg[key]['order_people'] += order_people
        agg[key]['impressions'] += impressions
        agg[key]['clicks'] += clicks
        if cpc > 0:
            agg[key]['cpc_sum'] += cpc
            agg[key]['cpc_count'] += 1
        if cpm > 0:
            agg[key]['cpm_sum'] += cpm
            agg[key]['cpm_count'] += 1
        if cvr > 0:
            agg[key]['cvr_sum'] += cvr
            agg[key]['cvr_count'] += 1
    
    print(f"\n=== 目标类型映射 ===")
    for k, v in sorted(target_names.items()):
        print(f"  索引 {k}")
    
    print(f"\n=== 城市总数: {len(cities)} ===")
    print(f"前30个城市: {sorted(cities)[:30]}")
    print(f"最后10个: {sorted(cities)[-10:]}")
    
    print(f"\n=== 聚合数据摘要（所有目标类型×城市组合）===")
    print(f"总组合数: {len(agg)}")
    
    # Show summary per target type
    for tt in sorted(agg.keys()):
        t_type = tt[0]
        city = tt[1]
        d = agg[tt]
        wroi = d['roi_sum'] / d['roi_count'] if d['roi_count'] > 0 else 0
        avg_cpc = d['cpc_sum'] / d['cpc_count'] if d['cpc_count'] > 0 else 0
        avg_cpm = d['cpm_sum'] / d['cpm_count'] if d['cpm_count'] > 0 else 0
        avg_cvr = d['cvr_sum'] / d['cvr_count'] if d['cvr_count'] > 0 else 0
        print(f"\n  {t_type} | {city}")
        print(f"    消耗(CA): {d['consumption']:.2f} | 注册会员: {d['new_members']} | GMV: {d['gmv']:.2f}")
        print(f"    ROI(w): {wroi:.2f} | 订单: {d['orders']} | CPC: {avg_cpc:.2f} | CPM: {avg_cpm:.2f} | CVR: {avg_cvr:.4f}")
    
    print("\n=== 总览汇总 ===")
    total_consumption = sum(d['consumption'] for d in agg.values())
    total_members = sum(d['new_members'] for d in agg.values())
    total_gmv = sum(d['gmv'] for d in agg.values())
    total_roi_w = sum(d['roi_sum'] for d in agg.values())
    total_roi_w_count = sum(d['roi_count'] for d in agg.values())
    print(f"总消耗(CA): {total_consumption:.2f}")
    print(f"总注册会员: {total_members}")
    print(f"总GMV: {total_gmv:.2f}")
    print(f"综合ROI(w): {total_roi_w/total_roi_w_count:.2f}" if total_roi_w_count > 0 else "N/A")
    
    print("\n=== 补充：首销会员数字段 ===")
    # Check if there's a "首销会员数" field in shared strings
    for i, s in enumerate(strings):
        if '首销' in s or '首次' in s or '首单' in s:
            print(f"  索引{i}: {s}")
    
    # Also check top 10 rows for any extra fields
    if len(strings) > 26:
        print(f"Shared strings 26-40:")
        for i, s in enumerate(strings[26:40]):
            print(f"  [{i+26}] {s[:60]}")
    
    # Print some data for later charting
    print("\n=== 省份级别数据（按消耗排序）===")
    # Aggregate by city only (combined across target types)
    city_agg = defaultdict(lambda: {'consumption': 0, 'members': 0, 'gmv': 0, 'roi_w': 0, 'roi_n': 0, 'orders': 0, 'clicks': 0, 'impressions': 0})
    for (tt, city), d in agg.items():
        city_agg[city]['consumption'] += d['consumption']
        city_agg[city]['members'] += d['new_members']
        city_agg[city]['gmv'] += d['gmv']
        city_agg[city]['roi_w'] += d['roi_sum']
        city_agg[city]['roi_n'] += d['roi_count']
        city_agg[city]['orders'] += d['orders']
        city_agg[city]['clicks'] += d['clicks']
        city_agg[city]['impressions'] += d['impressions']
    
    sorted_cities = sorted(city_agg.items(), key=lambda x: x[1]['consumption'], reverse=True)
    print(f"\nTop 40 cities by consumption:")
    for city_name, data in sorted_cities[:40]:
        wroi = data['roi_w'] / data['roi_n'] if data['roi_n'] > 0 else 0
        print(f"  {city_name:10s} | CA={data['consumption']:>10.2f} | 注册={data['members']:>6d} | GMV={data['gmv']:>10.2f} | ROI={wroi:>6.2f} | 订单={data['orders']:>5d}")
