#!/usr/bin/env python3
"""读取品牌会员数据.xlsx 的结构：sheet名 + 列头 + 前5行"""
import zipfile, re, os

xlsx_path = "/Users/edy/Desktop/品牌会员数据.xlsx"
out_path = "/Users/edy/Desktop/member_structure.txt"

# 1. 读 workbook.xml 拿 sheet 名
with zipfile.ZipFile(xlsx_path, 'r') as z:
    with z.open('xl/workbook.xml') as f:
        content = f.read().decode()
        sheets = re.findall(r'<sheet[^>]*name="([^"]+)"', content)
    
    # 2. 读 sheet1.xml 前几行拿列头
    with z.open('xl/worksheets/sheet1.xml') as f:
        raw = f.read().decode()
        # 找到所有 <row> 中第一个 row 的 <c> 标签，提取列头
        # 先找到第一个 <row>...</row>
        rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
        
        header_cells = re.findall(r'<c[^>]*>(.*?)</c>', rows[0], re.DOTALL)
        headers = []
        for cell in header_cells:
            v = re.search(r'<v>(.*?)</v>', cell)
            if v:
                headers.append(v.group(1))
            else:
                # 可能是内联字符串
                is_ = re.search(r'<is>(.*?)</is>', cell)
                if is_:
                    t = re.search(r'<t[^>]*>(.*?)</t>', is_.group(1))
                    headers.append(t.group(1) if t else '')
                else:
                    headers.append('')
        
        # 取前5行原始数据
        sample_rows = []
        for row_xml in rows[1:6]:
            cells = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
            vals = []
            for cell in cells:
                v = re.search(r'<v>(.*?)</v>', cell)
                if v:
                    vals.append(v.group(1)[:60])
                else:
                    is_ = re.search(r'<is>(.*?)</is>', cell)
                    if is_:
                        t = re.search(r'<t[^>]*>(.*?)</t>', is_.group(1))
                        vals.append(t.group(1)[:60] if t else '')
                    else:
                        vals.append('')
            sample_rows.append(vals)

with open(out_path, 'w') as f:
    f.write(f"Sheet names: {sheets}\n\n")
    f.write(f"Total rows in sheet1: {len(rows)}\n")
    f.write(f"Columns ({len(headers)}): {headers}\n\n")
    f.write("=== First 5 rows ===\n")
    for i, row in enumerate(sample_rows):
        f.write(f"Row {i+2}: {row}\n")

print(f"Written to {out_path}")
