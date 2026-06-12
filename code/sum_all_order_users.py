#!/usr/bin/env python3
"""汇总所有行（不管TARGET）的广告订单人数（col21）"""
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

col21_total = 0.0
col1_targets = {}
row_count = 0

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

    # 广告订单人数是 col21 (0-indexed)
    col21 = row.get(21, 0)
    if col21 is None or col21 == 'NULL':
        col21 = 0
    try:
        col21 = float(col21)
    except:
        col21 = 0

    col21_total += col21

    # 记录目标类型
    target = str(row.get(1, '')).strip()
    if target not in col1_targets:
        col1_targets[target] = {'rows': 0, 'col21_sum': 0.0}
    col1_targets[target]['rows'] += 1
    col1_targets[target]['col21_sum'] += col21

    row_count += 1

print(f"总行数（含header）: {len(rows)}")
print(f"数据行数: {row_count}")
print(f"\n========== 广告订单人数（col21）汇总 ==========")
print(f"所有TARGET合计: {col21_total:,.0f}")
print(f"好哥说的数值:   1,290,728")
print(f"差异:          {col21_total - 1290728:+,.0f}")

print(f"\n========== 各TARGET类型明细 ==========")
for t, info in sorted(col1_targets.items(), key=lambda x: x[1]['col21_sum'], reverse=True):
    label = f"'{t}'" if t else "(空)"
    print(f"  {label:30s} | 行数: {info['rows']:>8,d} | 广告订单人数: {info['col21_sum']:>12,.0f}")

print(f"\n{'='*60}")
print(f"总计: {col21_total:,.0f}")
