#!/usr/bin/env python3
"""深度诊断：检查注册会员列的真实映射关系"""
import zipfile, re
from collections import defaultdict

def col_letter_to_index(col_str):
    result = 0
    for c in col_str:
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result - 1

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = __import__('xml.etree.ElementTree', fromlist=['ElementTree']).ElementTree
    # 直接从文件解析
    import xml.etree.ElementTree as ET
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # 检查 K列值为 30 的行（不是共享字符串，是数值30）的上下文
    print("=== K列数值=30 的行样本 ===")
    count_30 = 0
    for idx, row_xml in enumerate(rows[1:]):
        k_cell = re.search(r'<c r="K(\d+)"[^>]*>(.*?)</c>', row_xml)
        if k_cell:
            row_num = k_cell.group(1)
            content = k_cell.group(2)
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm and vm.group(1) == '30' and 't="s"' not in content:
                # 解析整行
                all_cells = re.findall(r'<c[^>]*r="([^"]+)"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
                row_data = {}
                for ref, c in all_cells:
                    col = col_letter_to_index(re.match(r'([A-Z]+)', ref).group(1))
                    v = re.search(r'<v>(.*?)</v>', c)
                    if v:
                        val = v.group(1)
                        if 't="s"' in c:
                            try:
                                val = strings[int(val)]
                            except:
                                pass
                        row_data[col] = val
                
                if count_30 < 10:
                    print(f"  Row {row_num}: date={row_data.get(2)} city={row_data.get(5)} target={row_data.get(1)} plan={row_data.get(3)} cost={row_data.get(9)} reg_val=30(数值), gmv={row_data.get(24)}")
                count_30 += 1
    
    print(f"\nK列数值=30 的总行数: {count_30}")
    
    # 统计K列各种值的分布（不含shared string）
    print("\n=== K列数值分布 ===")
    val_dist = defaultdict(int)
    for row_xml in rows[1:]:
        k_cell = re.search(r'<c r="K(\d+)"[^>]*>(.*?)</c>', row_xml)
        if k_cell:
            content = k_cell.group(2)
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm and 't="s"' not in content:
                val_dist[vm.group(1)] += 1
    
    for v, cnt in sorted(val_dist.items(), key=lambda x: -x[1])[:15]:
        print(f"  数值 {v}: {cnt}行")
    
    # 没有K列的行数
    no_k = sum(1 for row_xml in rows[1:] if not re.search(r'<c r="K\d+"', row_xml))
    print(f"\n没有K列的行数: {no_k}")
    
    # 检查注册会员的真实列 — 数据里的K列并不一定对应注册会员，因为列和注册会员是通过字符串索引0,1,2...映射的
    # header行中K1的值是字符串索引10，对应shared strings[10]="注册会员人数（直接）"
    # 但数据行中K2的值如果是数值30，那30就是实际注册人数
    # 而 K4 的值如果是 t="s" + v=30，30是字符串索引->"NULL"
    
    # 重点：实际检查一个已知有注册会员的行
    # 从之前输出看 Row 9 (重庆, target=27) 的 reg=1
    # 验证一下
    row9 = rows[9]  # index 9 = 第9个数据行
    k9 = re.search(r'<c r="K10"[^>]*>(.*?)</c>', row9)
    if k9:
        print(f"\nRow 10 (重庆, reg=1) K10: {k9.group()}")
    else:
        # 找重庆那行
        for idx, row_xml in enumerate(rows[1:20]):
            if '重庆' in row_xml:
                print(f"\nRow {idx+2} (重庆):")
                k_cell = re.search(r'<c r="K\d+"[^>]*>(.*?)</c>', row_xml)
                if k_cell:
                    print(f"  K cell: {k_cell.group()}")
                else:
                    print(f"  No K cell found")
                    # 找所有cell
                    all_c = re.findall(r'<c[^>]*>(.*?)</c>', row_xml, re.DOTALL)
                    print(f"  All cells count: {len(all_c)}")
                    # K列应该是第11个cell(0-indexed=10)
                    if len(all_c) > 10:
                        print(f"  All cells: {[c[:80] for c in all_c]}")
                break
