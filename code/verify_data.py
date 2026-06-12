#!/usr/bin/env python3
import json

with open('/Users/edy/Desktop/看板代码/dashboard_data.json') as f:
    d = json.load(f)

acq_ou = sum(d['acq_order_users'])
ret_ou = sum(d['ret_order_users'])
total = acq_ou + ret_ou

print(f'拉新广告订单人数: {acq_ou:,}')
print(f'老客转化广告订单人数: {ret_ou:,}')
print(f'合计: {total:,}')
print(f'目标值:  1,290,728')
print(f'差异: {1290728 - total:,}')
print()
print('是否完全对上:', '✅ YES' if total == 1290728 else '❌ NO')
