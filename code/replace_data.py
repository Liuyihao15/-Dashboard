#!/usr/bin/env python3
"""Replace D in HTML with fresh JSON from dashboard_data.json"""
import json, re

with open('/Users/edy/Desktop/看板代码/dashboard_data.json', 'r', encoding='utf-8') as f:
    new_json = json.dumps(json.load(f), ensure_ascii=False)

with open('/Users/edy/Desktop/看板代码/brand_member_dashboard_v2.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Replace const D = {...} with new data
html = re.sub(
    r'const D = \{.*?\};',
    f'const D = {new_json};',
    html,
    flags=re.DOTALL
)

with open('/Users/edy/Desktop/看板代码/brand_member_dashboard_v2.html', 'w', encoding='utf-8') as f:
    f.write(html)

# Verify
verify = re.search(r'const D = (\{.*?\});', html, re.DOTALL)
if verify:
    d = json.loads(verify.group(1))
    has_trade = 'acq_trade_members' in d
    has_order = 'acq_order_users' in d
    print(f"OK: has_trade={has_trade}, has_order={has_order}")
    if has_trade:
        print(f"  trade_members total: {sum(d['acq_trade_members'])}")
    if has_order:
        print(f"  order_users total: {sum(d['acq_order_users'])}")
