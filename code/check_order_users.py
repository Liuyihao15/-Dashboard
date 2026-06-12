#!/usr/bin/env python3
import json

with open('/Users/edy/Desktop/看板代码/dashboard_data.json', 'r') as f:
    data = json.load(f)

total_order_users = sum(data['acq_order_users'])
print(f'下单用户数（修正后，列22→广告订单人数）总和: {total_order_users:,}')

print(f'\n每日明细:')
for i, d in enumerate(data['dates']):
    print(f'  {d}: {data["acq_order_users"][i]:,}人')

print(f'\n好哥说的广告订单人数: 1,290,728')
print(f'差异: {1290728 - total_order_users:,}')
