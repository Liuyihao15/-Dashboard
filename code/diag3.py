#!/usr/bin/env python3
"""验证：K列(注册会员)的值到底是数字还是 shared string 索引"""
import zipfile, re
from xml.etree import ElementTree as ET

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    print(f"Shared strings count: {len(strings)}")
    print(f"String[30] = '{strings[30]}'")
    print(f"String[32] = '{strings[32]}'")
    print(f"String[0] = '{strings[0]}'")
    
    # 看K列(K2)的XML
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    
    # Row 2 (第一个数据行) - reg=0
    row2_xml = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)[1]
    # 找到 K2 的cell
    k2 = re.search(r'<c r="K2"[^>]*>.*?</c>', row2_xml)
    print(f"\nK2 XML: {k2.group() if k2 else 'NOT FOUND'}")
    
    # Row 4 (第3个数据行) - reg看起来是30
    row4_xml = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)[3]
    k4 = re.search(r'<c r="K4"[^>]*>.*?</c>', row4_xml)
    print(f"K4 XML: {k4.group() if k4 else 'NOT FOUND'}")
    
    # Row 12 (第11个数据行) - reg也是30
    row12_xml = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)[11]
    k12 = re.search(r'<c r="K12"[^>]*>.*?</c>', row12_xml)
    print(f"K12 XML: {k12.group() if k12 else 'NOT FOUND'}")
    
    # 看看目标类型=32的行中，K列的值分布
    print("\n分析 target=32(复购) 的行中K列的值分布...")
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    k_vals_counter = {}
    for row_xml in rows[1:]:
        cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        row = {}
        for ref, content in cells:
            col_idx = None
            m = re.match(r'([A-Z]+)', ref)
            if m:
                col_str = m.group(1)
                col_idx = 0
                for c in col_str:
                    col_idx = col_idx * 26 + (ord(c) - ord('A') + 1)
                col_idx -= 1
            val = None
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm:
                val = vm.group(1)
                if 't="s"' in content:
                    try:
                        idx = int(val)
                        val = strings[idx] if idx < len(strings) else val
                    except:
                        pass
            row[col_idx] = val
        
        target = row.get(1)
        k_val = row.get(10)
        if target and str(target) == '32' and k_val:
            k_vals_counter[k_val] = k_vals_counter.get(k_val, 0) + 1
    
    print(f"K列在target=32行中的值分布: {k_vals_counter}")
