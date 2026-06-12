#!/usr/bin/env python3
"""Generate dashboard HTML with data injection"""
import json

with open('/Users/edy/Desktop/dashboard_data.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

json_data = json.dumps(data, ensure_ascii=False)

# Read the HTML template (no f-strings, no curly braces issue)
html_content = open('/Users/edy/Desktop/dashboard_v2.template.html', 'r', encoding='utf-8').read()

# Replace data placeholder
html_content = html_content.replace('__DATA_JSON__', json_data)

with open('/Users/edy/Desktop/brand_member_dashboard_v2.html', 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Generated: {len(html_content)} bytes")
