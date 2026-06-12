#!/usr/bin/env python3
"""诊断：看第一行非头数据的cell结构"""
import zipfile, re

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # 第2行（数据行）
    row2 = rows[1]
    print("=== Row 2 (first data row) full XML ===")
    print(row2[:3000])
    print()
    
    # 检查第1行是不是header
    row1 = rows[0]
    print("=== Row 1 header full XML ===")
    print(row1[:1000])
