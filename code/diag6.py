#!/usr/bin/env python3
"""最终诊断：Row 12 K列的精确XML"""
import zipfile, re

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # Row 12 (索引11)
    row12 = rows[11]
    print("=== Row 12 完整XML ===")
    print(row12)
    print()
    
    # 找到K列的cell
    k_match = re.search(r'(<c r="K\d+"[^>]*>.*?</c>)', row12)
    if k_match:
        print("K cell:", k_match.group(1))
        if 't="s"' in k_match.group(1):
            print("→ 包含 t='s' (shared string)")
        else:
            print("→ 不包含 t='s' (数值)")
    
    # 同样的方法看注册会员实际含义
    # 看看有没有注册会员人数>0的正常行
    # 第8行（重庆）reg=1
    row8 = rows[7]  # 第8个数据行 = index 7
    k8_match = re.search(r'(<c r="K\d+"[^>]*>.*?</c>)', row8)
    print("\n=== Row 8 (重庆, reg=1) ===")
    print(f"K cell: {k8_match.group(1) if k8_match else 'NOT FOUND'}")
    if k8_match:
        has_ts = 't="s"' in k8_match.group(1)
        print(f"Has t='s': {has_ts}")
    
    # 自己有一个注册会员的普通行 — 找K列值=3的行
    for i, row_xml in enumerate(rows[1:51]):
        k_cell = re.search(r'<c r="K(\d+)"[^>]*>(.*?)</c>', row_xml)
        if k_cell:
            content = k_cell.group(2)
            vm = re.search(r'<v>(.*?)</v>', content)
            k_num = k_cell.group(1)
            check_ts = 't="s"' in content
            if vm and vm.group(1) == '3' and not check_ts:
                print(f"\n=== Row {k_num} (reg=3) K cell XML ===")
                print(f"Full: {k_cell.group()}")
                # 解析整行
                all_cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
                for ref, c in all_cells:
                    v = re.search(r'<v>(.*?)</v>', c)
                    if v:
                        print(f"  {ref}: val={v.group(1)}, t_s={'t=\"s\"' in c}")
                break
