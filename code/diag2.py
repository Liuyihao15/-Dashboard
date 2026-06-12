#!/usr/bin/env python3
"""诊断：查看注册会员人数列和消耗列在各行的具体值，对比是否一致"""
import zipfile, re
from xml.etree import ElementTree as ET

def col_letter_to_index(col_str):
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

def parse_cell_value(cell_xml, strings):
    m = re.search(r'<v>(.*?)</v>', cell_xml)
    if not m:
        return None
    val = m.group(1)
    if 't="s"' in cell_xml or 't="inlineStr"' in cell_xml:
        try:
            idx = int(val)
            return strings[idx] if idx < len(strings) else val
        except:
            return val
    return val

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    print("=== 列名索引确认 ===")
    for i, name in enumerate(strings[:26]):
        print(f"  {i:2d}={name}")
    
    # 列索引：K=会员注册=10, J=消耗=9
    # 在A1表示中：A=0, B=1, ..., K=10
    print("\n目标索引: 注册会员人数（直接）-> col 10 = K列")
    print("目标索引: 广告消耗 -> col 9 = J列")
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    
    # 看第一行header的结构
    print("\n=== 第一行(header) cell引用 ===")
    first_row = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)[0]
    cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', first_row, re.DOTALL)
    for ref, content in cells:
        col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
        val = parse_cell_value(f'<c>{content}</c>', strings)
        print(f"  {ref} -> col {col} -> val='{val}' (string index)")
    
    # 看第二行（第一个数据行）
    print("\n=== 第二行(第一个数据行) cell引用 ===")
    second_row = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)[1]
    cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', second_row, re.DOTALL)
    for ref, content in cells:
        col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
        val = parse_cell_value(f'<c>{content}</c>', strings)
        print(f"  {ref} -> col {col} -> val='{val}'")
    
    # 检查 K列（注册会员）和 J列（消耗）在所有行中的值分布
    print("\n=== K列(注册会员)前30行实际值 ===")
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    for i, row_xml in enumerate(rows[1:31]):
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        row = {}
        for ref, content in cells:
            col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
            val = parse_cell_value(f'<c>{content}</c>', strings)
            row[col] = val
        cost = row.get(9, 'N/A')
        reg = row.get(10, 'N/A')
        city = row.get(5, 'N/A')
        date = row.get(2, 'N/A')
        target = row.get(1, 'N/A')
        plan_id = row.get(3, 'N/A')
        print(f"  Row {i+2}: date={date} city={city} target={target} plan={plan_id} cost={cost} reg={reg}")
