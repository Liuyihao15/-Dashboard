#!/usr/bin/env python3
"""完整读取品牌会员数据.xlsx：列名映射 + 各类别样本 + 数据概览"""
import zipfile, os
from xml.etree import ElementTree as ET

xlsx_path = "/Users/edy/Desktop/品牌会员数据.xlsx"
out_path = "/Users/edy/Desktop/member_overview.txt"

import re

with zipfile.ZipFile(xlsx_path, 'r') as z:
    # 读取 sharedStrings
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = []
    for si in sis:
        t = si.find('.//s:t', ns)
        strings.append(t.text if t is not None else '')
    
    # 列名 = shared strings 的前26个
    columns = strings[:26]
    
    # 读 sheet1 完整解析
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows_xml = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    print(f"Total rows: {len(rows_xml)}", file=open(out_path, 'w'))
    print(f"\n=== Columns ({len(columns)}) ===", file=open(out_path, 'a'))
    for i, col in enumerate(columns):
        print(f"  {i:2d}. {col}", file=open(out_path, 'a'))
    
    # 取前10行数据（把 shared string 索引转成文本）
    print(f"\n=== First 10 rows (value preview) ===", file=open(out_path, 'a'))
    for row_idx, row_xml in enumerate(rows_xml[:11]):
        cells = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        vals = []
        for cell in cells:
            v = re.search(r'<v>(.*?)</v>', cell)
            if v:
                val = v.group(1)
                # 如果是共享字符串，转成文本
                if 't="s"' in cell or 't="inlineStr"' in cell:
                    try:
                        idx = int(val)
                        if idx < len(strings):
                            vals.append(strings[idx][:40])
                        else:
                            vals.append(val)
                    except:
                        vals.append(val[:40])
                else:
                    vals.append(val[:40])
            else:
                vals.append('')
        print(f"  Row {row_idx+1}: {vals}", file=open(out_path, 'a'))
    
    # 统计各维度的去重值（前10个）
    print(f"\n=== Dimension unique values (first 10) ===", file=open(out_path, 'a'))
    
    # 列索引映射
    dim_cols = {
        '业务品牌名称': 0,       # strings[0]
        '目标类型name': 1,       # strings[1]
        '日': 2,                 # strings[2] - 日期
        '省份': 5,               # strings[5]
        '城市': 6,               # strings[6]
    }
    
    for dim_name, col_idx in dim_cols.items():
        # 提取这一列的所有值
        col_vals = set()
        for row_xml in rows_xml[1:]:  # skip header
            cells = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
            if col_idx < len(cells):
                cell = cells[col_idx]
                v = re.search(r'<v>(.*?)</v>', cell)
                if v:
                    val = v.group(1)
                    if 't="s"' in cell:
                        try:
                            idx = int(val)
                            if idx < len(strings):
                                col_vals.add(strings[idx][:40])
                        except:
                            pass
                    else:
                        col_vals.add(val[:40])
        col_vals.discard('0')  # 去掉可能的空值
        sorted_vals = sorted(col_vals, key=lambda x: (x is None or x == 'NULL', x))
        print(f"\n  {dim_name}: {sorted_vals[:15]}", file=open(out_path, 'a'))
        print(f"  (total unique: {len(col_vals)})", file=open(out_path, 'a'))

print(f"Written to {out_path}")
