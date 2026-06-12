#!/usr/bin/env python3
"""Read brand member data xlsx - get column names and first 2 rows sample"""
import zipfile, re
from xml.etree import ElementTree as ET

path = '/Users/edy/Desktop/品牌会员数据.xlsx'

with zipfile.ZipFile(path, 'r') as z:
    # Read shared strings (column names + data)
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    # Read sheet1
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    
    # Get all rows (with or without attributes)
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    print(f"Total shared strings: {len(strings)}")
    print(f"Total rows in sheet: {len(rows)}")
    print()
    
    # Column names are first row
    first_row_cells = re.findall(r'<c[^>]*>(.*?)</c>', rows[0], re.DOTALL)
    columns = []
    for cell in first_row_cells:
        v = re.search(r'<v>(.*?)</v>', cell)
        t_attr = re.search(r't="(\w+)"', cell)
        if v:
            val = v.group(1)
            if t_attr and t_attr.group(1) == 's':
                idx = int(val)
                columns.append(strings[idx] if idx < len(strings) else val)
            else:
                columns.append(val)
        else:
            columns.append('')
    
    print(f"=== 列字段 ({len(columns)}列) ===")
    for i, c in enumerate(columns):
        print(f"  [{i:2d}] {c}")
    
    print()
    
    # Print first 2 data rows
    for row_idx in range(1, min(4, len(rows))):
        cells = re.findall(r'<c[^>]*>(.*?)</c>', rows[row_idx], re.DOTALL)
        vals = []
        for cell in cells:
            v = re.search(r'<v>(.*?)</v>', cell)
            t_attr = re.search(r't="(\w+)"', cell)
            ref = re.search(r'r="(\w+)"', cell)
            if v:
                val = v.group(1)
                if t_attr and t_attr.group(1) == 's':
                    idx = int(val)
                    vals.append(strings[idx] if idx < len(strings) else val)
                else:
                    vals.append(val)
            else:
                vals.append('')
        print(f"\nData Row {row_idx}:")
        for i, val in enumerate(vals):
            print(f"  [{i:2d}] {columns[i]:30s} = {val[:60]}")
