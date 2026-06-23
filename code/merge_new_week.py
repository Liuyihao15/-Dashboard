#!/usr/bin/env python3
"""Merge new week data (0608-0614) into existing dashboard_data.json"""
import json

# Load existing data
with open('/Users/edy/Desktop/霸王茶姬看板/dashboard_data.json') as f:
    existing = json.load(f)

# Load new week data
with open('/Users/edy/Desktop/看板代码/data_week_0608_0614.json') as f:
    new_data = json.load(f)

print(f"Existing dates: {existing['dates'][:3]}...{existing['dates'][-3:]}")
print(f"New week dates: {new_data['dates']}")

# --- 1. Merge daily time series ---
# Check for overlap
existing_dates = set(existing['dates'])
new_dates_set = set(new_data['dates'])
overlap = existing_dates & new_dates_set
if overlap:
    print(f"WARNING: Overlapping dates: {sorted(overlap)}")
else:
    print("No date overlap — safe to merge")

# Fields that map directly (same keys in both files)
daily_fields = ['acq_cost', 'acq_members', 'acq_gmv', 'acq_order_users',
                'ret_cost', 'ret_members', 'ret_order_users',
                'acq_orders', 'ret_gmv', 'ret_orders',
                'acq_trade_members']

# Extend existing arrays with new data (in date order)
all_dates = sorted(existing['dates'] + new_data['dates'])

# Build merged daily data
merged = {'dates': all_dates}

# For each time-series field, build complete array
for field in ['acq_cost', 'acq_members', 'acq_gmv', 'acq_order_users',
              'ret_cost', 'ret_members', 'ret_order_users',
              'acq_trade_members']:
    # Build dict from date->value
    existing_dict = dict(zip(existing['dates'], existing.get(field, [0]*len(existing['dates']))))
    new_dict = dict(zip(new_data['dates'], new_data.get(field, [0]*len(new_data['dates']))))
    merged_dict = {**existing_dict, **new_dict}
    merged[field] = [merged_dict[d] for d in all_dates]

# Fill any gaps for fields only in new data
for field in ['acq_orders']:
    if field in new_data and field not in existing:
        existing_dict = dict(zip(existing['dates'], [0]*len(existing['dates'])))
        new_dict = dict(zip(new_data['dates'], new_data.get(field, [0]*len(new_data['dates']))))
        merged_dict = {**existing_dict, **new_dict}
        merged[field] = [merged_dict[d] for d in all_dates]

# Also merge ret_gmv, ret_orders if present in new_data
for field in ['ret_gmv', 'ret_orders']:
    if field in new_data:
        existing_dict = dict(zip(existing['dates'], existing.get(field, [0]*len(existing['dates']))))
        new_dict = dict(zip(new_data['dates'], new_data.get(field, [0]*len(new_data['dates']))))
        merged_dict = {**existing_dict, **new_dict}
        merged[field] = [merged_dict[d] for d in all_dates]

# --- 2. Handle full_members (keeps track of total member base) ---
# If existing has full_members, extend with last value (no new data for this)
if 'full_members' in existing:
    merged['full_members'] = existing['full_members'] + [existing['full_members'][-1]] * len(new_data['dates'])

# --- 3. Province data needs recalculation ---
# The new week data has its own province breakdown.
# But the dashboard uses prov_cost/prov_members/prov_gmv as totals for the ENTIRE period.
# So we need to ADD the new week's province data to the existing.
# Actually, looking at the code more carefully - prov_cost/Members/GMV are CUMULATIVE totals
# for the date range shown. The data is from the same source file.
# 
# The new week data comes from a different xlsx export (only 6/8-6/14).
# To merge properly, we should ADD the province-level data.
# However, this is tricky because the province list might differ slightly.
# 
# The simplest approach: use the new week data to RECALCULATE province totals
# from the full date range. But we don't have daily-by-province data.
# 
# Alternative: since prov data is cumulative from the SAME xlsx export,
# and the new week data is from a SEPARATE export, we need to add them.
# 
# ACTUALLY: looking at the template more carefully, prov_cost etc are 
# simply the province-level aggregation computed once from the xlsx.
# Since we're merging two separate xlsx exports, we need to SUM the province data.

# Build combined province stats
province_names = sorted(set(existing['provinces'] + new_data['provinces']))

# Build province-level data by summing
prov_total_cost = {}
prov_total_members = {}
prov_total_gmv = {}

# Initialize from existing
for p, c, m, g in zip(existing['provinces'], existing['prov_cost'], 
                        existing['prov_members'], existing['prov_gmv']):
    prov_total_cost[p] = prov_total_cost.get(p, 0) + c
    prov_total_members[p] = prov_total_members.get(p, 0) + m
    prov_total_gmv[p] = prov_total_gmv.get(p, 0) + g

# Add new week data
for p, c, m, g in zip(new_data['provinces'], new_data['prov_cost'],
                        new_data['prov_members'], new_data['prov_gmv']):
    prov_total_cost[p] = prov_total_cost.get(p, 0) + c
    prov_total_members[p] = prov_total_members.get(p, 0) + m
    prov_total_gmv[p] = prov_total_gmv.get(p, 0) + g

# Sort by cost descending
province_sorted = sorted(province_names, key=lambda p: prov_total_cost[p], reverse=True)

merged['provinces'] = province_sorted
merged['prov_cost'] = [round(prov_total_cost[p], 2) for p in province_sorted]
merged['prov_members'] = [int(prov_total_members[p]) for p in province_sorted]
merged['prov_gmv'] = [round(prov_total_gmv[p], 2) for p in province_sorted]
merged['prov_cac'] = [round(prov_total_cost[p]/prov_total_members[p], 2) if prov_total_members[p]>0 else 0 for p in province_sorted]
merged['prov_roi'] = [round(prov_total_gmv[p]/prov_total_cost[p], 2) if prov_total_cost[p]>0 else 0 for p in province_sorted]

# --- 4. Province break-down (new/ret) ---
# Merge prov_new_cost, prov_new_members, prov_new_gmv
if all(k in new_data for k in ['prov_new_cost', 'prov_new_members', 'prov_new_gmv']):
    prov_new_cost = {}
    prov_new_members = {}
    prov_new_gmv = {}
    
    for p, c, m, g in zip(existing.get('provinces', []), 
                            existing.get('prov_new_cost', []),
                            existing.get('prov_new_members', []),
                            existing.get('prov_new_gmv', [])):
        prov_new_cost[p] = prov_new_cost.get(p, 0) + c
        prov_new_members[p] = prov_new_members.get(p, 0) + m
        prov_new_gmv[p] = prov_new_gmv.get(p, 0) + g
    
    for p, c, m, g in zip(new_data['provinces'], new_data['prov_new_cost'],
                            new_data['prov_new_members'], new_data['prov_new_gmv']):
        prov_new_cost[p] = prov_new_cost.get(p, 0) + c
        prov_new_members[p] = prov_new_members.get(p, 0) + m
        prov_new_gmv[p] = prov_new_gmv.get(p, 0) + g
    
    merged['prov_new_cost'] = [round(prov_new_cost.get(p, 0), 2) for p in province_sorted]
    merged['prov_new_members'] = [int(prov_new_members.get(p, 0)) for p in province_sorted]
    merged['prov_new_gmv'] = [round(prov_new_gmv.get(p, 0), 2) for p in province_sorted]

# Merge prov_ret_cost, prov_ret_members
if all(k in new_data for k in ['prov_ret_cost', 'prov_ret_members']):
    prov_ret_cost = {}
    prov_ret_members = {}
    
    for p, c, m in zip(existing.get('provinces', []),
                         existing.get('prov_ret_cost', []),
                         existing.get('prov_ret_members', [])):
        prov_ret_cost[p] = prov_ret_cost.get(p, 0) + c
        prov_ret_members[p] = prov_ret_members.get(p, 0) + m
    
    for p, c, m in zip(new_data['provinces'], new_data['prov_ret_cost'],
                         new_data['prov_ret_members']):
        prov_ret_cost[p] = prov_ret_cost.get(p, 0) + c
        prov_ret_members[p] = prov_ret_members.get(p, 0) + m
    
    merged['prov_ret_cost'] = [round(prov_ret_cost.get(p, 0), 2) for p in province_sorted]
    merged['prov_ret_members'] = [int(prov_ret_members.get(p, 0)) for p in province_sorted]

# --- 5. City data ---  
# Merge city stats
city_list = sorted(set(existing.get('cities', []) + new_data.get('cities', [])))
city_cost = {}
city_members = {}
city_gmv = {}

for c_name, c_val, m_val, g_val in zip(existing.get('cities', []),
                                         existing.get('city_cost', []),
                                         existing.get('city_members', []),
                                         existing.get('city_gmv', [])):
    city_cost[c_name] = city_cost.get(c_name, 0) + c_val
    city_members[c_name] = city_members.get(c_name, 0) + m_val
    city_gmv[c_name] = city_gmv.get(c_name, 0) + g_val

for c_name, c_val, m_val, g_val in zip(new_data.get('cities', []),
                                         new_data.get('city_cost', []),
                                         new_data.get('city_members', []),
                                         new_data.get('city_gmv', [])):
    city_cost[c_name] = city_cost.get(c_name, 0) + c_val
    city_members[c_name] = city_members.get(c_name, 0) + m_val
    city_gmv[c_name] = city_gmv.get(c_name, 0) + g_val

city_sorted = sorted(city_list, key=lambda c: city_cost.get(c, 0), reverse=True)

merged['cities'] = city_sorted
merged['city_cost'] = [round(city_cost.get(c, 0), 2) for c in city_sorted]
merged['city_members'] = [int(city_members.get(c, 0)) for c in city_sorted]
merged['city_gmv'] = [round(city_gmv.get(c, 0), 2) for c in city_sorted]
merged['city_roi'] = [round((city_gmv.get(c, 0))/(city_cost.get(c, 1) or 1), 2) for c in city_sorted]
merged['city_cac'] = [round((city_cost.get(c, 0))/(city_members.get(c, 1) or 1), 2) for c in city_sorted]

# City province mapping - preserve from both
city_prov = {}
for c_name, p in zip(existing.get('cities', []), existing.get('city_prov', [])):
    city_prov[c_name] = p
for c_name, p in zip(new_data.get('cities', []), new_data.get('city_prov', [])):
    city_prov[c_name] = p
merged['city_prov'] = [city_prov.get(c, '') for c in city_sorted]

# --- 6. Recalculate KPIs ---
total_cost = sum(merged['prov_cost'])
total_members = sum(merged['prov_members'])
total_gmv = sum(merged['prov_gmv'])
acq_total_cost = sum(merged.get('acq_cost', []))
acq_total_members = sum(merged.get('acq_members', []))

merged['kpi'] = {
    'total_cost': round(total_cost, 2),
    'total_members': total_members,
    'total_gmv': round(total_gmv, 2),
    'total_roi': round(total_gmv/total_cost, 2) if total_cost>0 else 0,
    'total_cac': round(total_cost/total_members, 2) if total_members>0 else 0,
    'acq_cost': round(acq_total_cost, 2),
    'acq_members': acq_total_members,
    'acq_cac': round(acq_total_cost/acq_total_members, 2) if acq_total_members>0 else 0,
}

# Write merged data
output_path = '/Users/edy/Desktop/霸王茶姬看板/dashboard_data.json'
with open(output_path, 'w', encoding='utf-8') as f:
    json.dump(merged, f, ensure_ascii=False, separators=(',',':'))

print(f"✅ Merged data written to {output_path}")
print(f"   Dates: {merged['dates'][0]} ~ {merged['dates'][-1]} ({len(merged['dates'])} days)")
print(f"   Provinces: {len(merged['provinces'])}")
print(f"   Cities: {len(merged['cities'])}")
print(f"   Total Cost: ¥{merged['kpi']['total_cost']/10000:.1f}万")
print(f"   Total Members: {merged['kpi']['total_members']:,}")
print(f"   Total GMV: ¥{merged['kpi']['total_gmv']/10000:.1f}万")
print(f"   Overall ROI: {merged['kpi']['total_roi']}")
print(f"   Overall CAC: ¥{merged['kpi']['total_cac']}")
