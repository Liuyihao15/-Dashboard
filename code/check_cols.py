#!/usr/bin/env python3
"""展示 shared strings 26-50 和列名"""
import zipfile, re
from xml.etree import ElementTree as ET

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = []
    for si in sis:
        t = si.find('.//s:t', ns)
        strings.append(t.text if t is not None else '')

    print('=== Columns ===')
    for i, c in enumerate(strings[:26]):
        print(f'{i:2d}: {c}')
    
    print()
    print('=== Shared strings 26-60 ===')
    for i, s in enumerate(strings[26:60]):
        print(f'{i+26:2d}: {s}')
    
    # 查一下目标类型name（索引1）的所有值
    print()
    print('=== All target types (col 1) ===')
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row>(.*?)</row>', raw, re.DOTALL)
    types = set()
    for row_xml in rows:
        cells = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        if len(cells) > 1:
            cell = cells[1]
            v = re.search(r'<v>(.*?)</v>', cell)
            if v:
                val = v.group(1)
                if 't="s"' in cell:
                    types.add(strings[int(val)])
                else:
                    types.add(val)
    types.discard('1')
    print(sorted(types))
