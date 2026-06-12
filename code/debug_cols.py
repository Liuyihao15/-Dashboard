#!/usr/bin/env python3
"""Debug: show shared strings and first few rows properly"""
import zipfile, re
from xml.etree import ElementTree as ET

path = '/Users/edy/Desktop/品牌会员数据.xlsx'

with zipfile.ZipFile(path, 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = []
    for si in sis:
        t = si.find('.//s:t', ns)
        strings.append(t.text if t is not None else '')
    
    print(f"Total shared strings: {len(sis)}")
    
    # Find the actual column names - look for common Chinese business terms
    print("\n=== All non-numeric shared strings ===")
    for i, s in enumerate(strings):
        if s and not s.replace('.', '').replace('-', '').isdigit():
            print(f"  [{i:3d}] {s[:80]}")
    
    # Read sheet1 - look at actual column names from the first row
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    
    # Parse the column names row (first row in sheet)
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    print(f"\nTotal rows: {len(rows)}")
    
    cell_re = re.compile(r'<c[^>]*>(.*?)</c>', re.DOTALL)
    
    # First row = column names
    first_cells = cell_re.findall(rows[0])
    print(f"\n=== Row 0 cells (column headers): {len(first_cells)} ===")
    for i, cell in enumerate(first_cells):
        ref = re.search(r'r="([^"]+)"', cell)
        t_attr = re.search(r't="(\w+)"', cell)
        v = re.search(r'<v>(.*?)</v>', cell)
        val = v.group(1) if v else ''
        if t_attr and t_attr.group(1) == 's':
            idx = int(val)
            print(f"  [{i:2d}] ref={ref.group(1) if ref else '?'} type=s idx={val} -> {strings[idx][:60]}")
        else:
            print(f"  [{i:2d}] ref={ref.group(1) if ref else '?'} type=inline val={val[:60]}")
    
    # Look at actual data rows (rows where col0 has non-zero value like 26)
    print("\n=== Scanning for actual data rows (col0=26 ≈ 霸王茶姬 brand ID) ===")
    data_rows = 0
    for row_idx, row_xml in enumerate(rows[1:6], 1):
        cells = cell_re.findall(row_xml)
        if not cells:
            continue
        # Check col0
        cell0 = cells[0]
        v0 = re.search(r'<v>(.*?)</v>', cell0)
        if v0:
            val0 = v0.group(1)
            t0 = re.search(r't="(\w+)"', cell0)
            if t0 and t0.group(1) == 's':
                val0 = strings[int(val0)]
            print(f"\nRow {row_idx}: col0={val0[:50]}, cells={len(cells)}")
            for i, cell in enumerate(cells):
                ref = re.search(r'r="([^"]+)"', cell)
                t_attr = re.search(r't="(\w+)"', cell)
                v = re.search(r'<v>(.*?)</v>', cell)
                if v:
                    val = v.group(1)
                    if t_attr and t_attr.group(1) == 's':
                        idx = int(val)
                        print(f"  [{i:2d}] {ref.group(1) if ref else '?'} type=s -> {strings[idx][:60]}")
                    else:
                        print(f"  [{i:2d}] {ref.group(1) if ref else '?'} type=num -> {val[:60]}")

    # Check what column 5 values look like for first 10 data rows
    print("\n=== Col 5 (city/province) for first 10 data rows ===")
    count = 0
    for row_xml in rows[1:]:
        cells = cell_re.findall(row_xml)
        if not cells:
            continue
        if len(cells) > 5:
            cell5 = cells[5]
            v = re.search(r'<v>(.*?)</v>', cell5)
            t_attr = re.search(r't="(\w+)"', cell5)
            if v:
                val = v.group(1)
                if t_attr and t_attr.group(1) == 's':
                    val = strings[int(val)]
                print(f"  col5={val[:50]}")
                count += 1
                if count >= 10:
                    break
