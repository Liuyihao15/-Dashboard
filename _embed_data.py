#!/usr/bin/env python3
"""Read dashboard_data.json, read template.html, embed data, write index.html"""
import json, sys

with open('/Users/edy/Desktop/霸王茶姬看板/dashboard_data.json') as f:
    data = json.load(f)

with open('/Users/edy/Desktop/霸王茶姬看板/template.html') as f:
    html = f.read()

# Embed data into __DATA__ placeholder
json_str = json.dumps(data, ensure_ascii=False, separators=(',',':'))
html = html.replace('__DATA__', json_str)

with open('/Users/edy/Desktop/霸王茶姬看板/index.html', 'w') as f:
    f.write(html)

print(f'✅ Embedded {len(json_str)} bytes of JSON data into index.html')
print(f'   Total index.html size: {len(html)} bytes')
