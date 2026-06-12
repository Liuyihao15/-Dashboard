#!/usr/bin/env python3
"""读取品牌会员数据.xlsx 的 shared strings 和列头"""
import zipfile, re, os
from xml.etree import ElementTree as ET

xlsx_path = "/Users/edy/Desktop/品牌会员数据.xlsx"
out_path = "/Users/edy/Desktop/member_columns.txt"

with zipfile.ZipFile(xlsx_path, 'r') as z:
    # 看有哪些文件
    all_files = [f.filename for f in z.filelist]
    print("All files:", all_files[:30], file=open(out_path, 'w'))
    
    # 读 sharedStrings
    has_ss = 'xl/sharedStrings.xml' in all_files
    print(f"\nHas sharedStrings: {has_ss}", file=open(out_path, 'a'))
    
    if has_ss:
        with z.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sis = root.findall('.//s:si', ns)
            strings = []
            for si in sis:
                t = si.find('.//s:t', ns)
                strings.append(t.text if t is not None else '')
            print(f"\nShared strings count: {len(strings)}", file=open(out_path, 'a'))
            print(f"First 50 strings: {strings[:50]}", file=open(out_path, 'a'))
    else:
        # 没有 sharedStrings，说明列头可能是内联字符串或数字
        # 看一下 styles.xml 了解数字格式
        print("\nNo sharedStrings - checking styles.xml", file=open(out_path, 'a'))
    
    # 读 sheet1.xml 第一行获取列引用
    with z.open('xl/worksheets/sheet1.xml') as f:
        raw = f.read().decode()
        # 找第一个 row 的所有 cell 引用
        rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
        first_row = rows[0]
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', first_row, re.DOTALL)
        
        print(f"\nFirst row cell refs (A1 notation):", file=open(out_path, 'a'))
        for ref, content in cells[:30]:
            # 提取值
            v = re.search(r'<v>(.*?)</v>', content)
            val = v.group(1) if v else ''
            # 提取是否内联字符串
            is_ = 't="inlineStr"' in content or 't="s"' in content
            print(f"  {ref}: val='{val[:60]}' inline={is_}", file=open(out_path, 'a'))

print(f"Done. Written to {out_path}")
