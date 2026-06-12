#!/usr/bin/env python3
"""修复数据提取：用正确的方式判断t=s"""
import zipfile, re
from xml.etree import ElementTree as ET
from collections import defaultdict

def col_letter_to_index(col_str):
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # 用正确的正则：<c ... ><v>...</v></c> 捕获cell的属性+内容
    # 方式：捕获整个 <c ... >...</c>
    
    total_reg = 0
    total_cost = 0
    
    for idx, row_xml in enumerate(rows[1:]):
        # 捕获整个cell
        cells = re.findall(r'(<c[^>]*>.*?</c>)', row_xml, re.DOTALL)
        row = {}
        for cell_str in cells:
            # 提取r属性
            rm = re.search(r'r="([^"]+)"', cell_str)
            if not rm:
                continue
            ref = rm.group(1)
            col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
            
            # 检查是否有t="s"
            is_str = 't="s"' in cell_str
            
            # 提取值
            vm = re.search(r'<v>(.*?)</v>', cell_str)
            if vm:
                val = vm.group(1)
                if is_str:
                    # 共享字符串索引
                    try:
                        val = strings[int(val)]
                    except:
                        pass
                row[col] = val
        
        # 累加
        k_val = row.get(10)
        if k_val and k_val not in ('NULL', 'None', ''):
            try:
                total_reg += int(float(k_val))
            except:
                pass
        
        j_val = row.get(9)
        if j_val:
            try:
                total_cost += float(j_val)
            except:
                pass
    
    print(f"正确累积的注册会员数: {total_reg}")
    print(f"正确累积的消耗: {total_cost:.2f}")
    
    # 输出一个样本行确认
    print("\n=== 确认: Row 12 K列 ===")
    row12_xml = rows[11]
    cells = re.findall(r'(<c[^>]*>.*?</c>)', row12_xml, re.DOTALL)
    for cell_str in cells:
        if 'K12' in cell_str:
            is_str = 't="s"' in cell_str
            vm = re.search(r'<v>(.*?)</v>', cell_str)
            val = vm.group(1) if vm else 'N/A'
            if is_str:
                try:
                    val = f"'{strings[int(val)]}' (was index {val})"
                except:
                    pass
            print(f"  K12: is_str={is_str}, val={val}")
    
    # 确认重庆行
    print("\n=== 确认: Row 8 (重庆) K列 ===")
    row8_xml = rows[7]
    cells = re.findall(r'(<c[^>]*>.*?</c>)', row8_xml, re.DOTALL)
    for cell_str in cells:
        if 'K8' in cell_str:
            is_str = 't="s"' in cell_str
            vm = re.search(r'<v>(.*?)</v>', cell_str)
            val = vm.group(1) if vm else 'N/A'
            print(f"  K8: is_str={is_str}, val={val}")
