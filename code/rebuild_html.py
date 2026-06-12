#!/usr/bin/env python3
"""Rebuild brand_member_dashboard_v2.html with latest dashboard_data.json embedded."""
import json, re

HTML_PATH = '/Users/edy/Desktop/看板代码/brand_member_dashboard_v2.html'
JSON_PATH = '/Users/edy/Desktop/看板代码/dashboard_data.json'

with open(HTML_PATH, 'r') as f:
    html = f.read()

with open(JSON_PATH, 'r') as f:
    data = json.load(f)

# Replace the inline D = {...} with new data
old = re.search(r'const D = \{.*?\};', html, re.DOTALL).group(0)
new = f'const D = {json.dumps(data, ensure_ascii=False, separators=[",",":"])};'
html = html.replace(old, new)

with open(HTML_PATH, 'w') as f:
    f.write(html)

print("✅ HTML updated with latest dashboard_data.json embedded")
print(f"   Dates: {len(data['dates'])} days")
print(f"   Provinces: {len(data['provinces'])}")
print(f"   Has ret_order_users: {'ret_order_users' in data}")
