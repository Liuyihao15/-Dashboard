#!/usr/bin/env python3
"""诊断数据准确性问题：逐行累加注册会员"""
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
    
    total_reg = 0
    total_cost = 0
    reg_s_values = set()
    reg_n_values = set()
    
    for idx, row_xml in enumerate(rows[1:]):
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        row = {}
        for ref, content in cells:
            col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm:
                val = vm.group(1)
                is_str = 't="s"' in content
                if is_str:
                    try:
                        idx_s = int(val)
                        val = strings[idx_s] if idx_s < len(strings) else val
                    except:
                        pass
                row[col] = (val, is_str)
            else:
                row[col] = (None, False)
        
        k_cell = row.get(10)
        if k_cell:
            k_val, is_str = k_cell
            if is_str:
                # shared string index - 可能是 'NULL'
                if k_val and k_val not in ('NULL', 'None', ''):
                    reg_s_values.add(k_val)
                # 不计入
            else:
                # 实际数字
                if k_val:
                    try:
                        v = int(float(k_val))
                        total_reg += v
                        reg_n_values.add(v)
                    except:
                        pass
        
        j_cell = row.get(9)
        if j_cell:
            j_val, is_str = j_cell
            if not is_str and j_val:
                try:
                    total_cost += float(j_val)
                except:
                    pass
    
    print(f"累计注册会员（仅非shared string的值）: {total_reg}")
    print(f"累计消耗: {total_cost}")
    print(f"注册会员非shared string值的种类: {sorted(reg_n_values)[:20]}")
    print(f"注册会员shared string值的种类: {reg_s_values}")
    print(f"注册会员shared string值行数 (NULL): {sum(1 for r in rows[1:] if True)}")
    
    # 检查是否有 K 列是 shared string 且值为其他内容（不是NULL）
    k_other_vals = defaultdict(int)
    for row_xml in rows[1:]:
        k_cell = re.search(r'<c r="K\d+"[^>]*>(.*?)</c>', row_xml)
        if k_cell:
            content = k_cell.group(1)
            if 't="s"' in content:
                vm = re.search(r'<v>(.*?)</v>', content)
                if vm:
                    try:
                        idx = int(vm.group(1))
                        s_val = strings[idx] if idx < len(strings) else vm.group(1)
                        k_other_vals[s_val] += 1
                    except:
                        pass
    
    print(f"\nK列 shared string 值分布:")
    for v, cnt in sorted(k_other_vals.items(), key=lambda x: -x[1]):
        print(f"  '{v}': {cnt}行")
