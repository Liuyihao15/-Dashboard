#!/usr/bin/env python3
"""检查 Row 12 的K列和注册会员列的真实XML和值"""
import zipfile, re

with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)
    
    # Row 12 (索引11)
    row12 = rows[11]
    print("=== Row 12 完整XML ===")
    print(row12[:500])
    print()
    
    # K12
    km = re.search(r'<c r="K12"[^>]*>.*?</c>', row12, re.DOTALL)
    if km:
        k12 = km.group()
        print("K12 XML:", k12)
        has_ts = 't="s"' in k12
        print("Has t=s:", has_ts)
        vm = re.search(r'<v>(.*?)</v>', k12)
        if vm:
            print("Value:", vm.group(1))
    else:
        print("K12 NOT FOUND via c r=K12")
        # 看所有K列
        all_k = re.findall(r'<c r="K\d+"[^>]*>.*?</c>', row12, re.DOTALL)
        print("All K cells:", all_k)
    
    print()
    print("=== Row 8 (重庆) ===")
    row8 = rows[7]
    k8m = re.search(r'<c r="K8"[^>]*>.*?</c>', row8, re.DOTALL)
    if k8m:
        k8 = k8m.group()
        print("K8:", k8)
        has_ts = 't="s"' in k8
        print("Has t=s:", has_ts)
        vm = re.search(r'<v>(.*?)</v>', k8)
        if vm:
            print("Value:", vm.group(1))
    
    print()
    print("=== 检查 K 列中值=30且带 t=s 的行 ===")
    count_ts_30 = 0
    count_no_ts_30 = 0
    for row_xml in rows[1:]:
        km = re.search(r'<c r="K\d+"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        if km:
            content = km.group(1)
            has_ts = 't="s"' in content
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm and vm.group(1) == '30':
                if has_ts:
                    count_ts_30 += 1
                else:
                    count_no_ts_30 += 1
    
    print("带 t=s 且值=30 的行数:", count_ts_30)
    print("不带 t=s 且值=30 的行数:", count_no_ts_30)
    
    # 检查带 t=s 的行中值=30的 — 这些的30是字符串索引
    val_dist_ts = {}
    val_dist_nots = {}
    for row_xml in rows[1:]:
        km = re.search(r'<c r="K\d+"[^>]*>(.*?)</c>', row_xml, re.DOTALL)
        if km:
            content = km.group(1)
            has_ts = 't="s"' in content
            vm = re.search(r'<v>(.*?)</v>', content)
            if vm:
                v = vm.group(1)
                if has_ts:
                    val_dist_ts[v] = val_dist_ts.get(v, 0) + 1
                else:
                    val_dist_nots[v] = val_dist_nots.get(v, 0) + 1
    
    print("\nK列带 t=s 的值分布:")
    for v, cnt in sorted(val_dist_ts.items(), key=lambda x: -x[1])[:10]:
        print(f"  val={v}: {cnt}行")
    print("\nK列不带 t=s 的值分布:")
    for v, cnt in sorted(val_dist_nots.items(), key=lambda x: -x[1])[:10]:
        print(f"  val={v}: {cnt}行")
