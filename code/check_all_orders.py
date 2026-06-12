#!/usr/bin/env python3
import zipfile, re
from xml.etree import ElementTree as ET

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)

print(f'总行数（含表头）: {len(rows)}')

total_order_users = 0
total_all_target = 0  # 所有行总广告订单人数（不含表头）
target_counts = {}   # 各TARGET的订单人数

for row_xml in rows[1:]:
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
            if val == 'NULL':
                row[col] = 0
            else:
                try: row[col] = float(val) if '.' in str(val) or not is_str else val
                except: row[col] = val

    # 列22 = 广告订单人数（W列）
    ou = float(row.get(22, 0) or 0)
    total_all_target += ou

    # 按目标类型汇总
    target = str(row.get(1, '')).strip()
    if target not in target_counts:
        target_counts[target] = {'rows': 0, 'order_users': 0}
    target_counts[target]['rows'] += 1
    target_counts[target]['order_users'] += ou

print(f'\n所有TARGET的广告订单人数合计: {int(total_all_target):,}')
print(f'\n各TARGET明细:')
for t, v in sorted(target_counts.items(), key=lambda x: -x[1]['order_users']):
    print(f'  TARGET="{t}": {v["rows"]}行, 广告订单人数={int(v["order_users"]):,}')
