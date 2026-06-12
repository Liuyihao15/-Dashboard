#!/usr/bin/env python3
"""Correct parsing — column 5 is '省份'(all NaN), column 6 is '城市'(actual city names)"""
import pandas as pd

path = '/Users/edy/Desktop/品牌会员数据.xlsx'

# Read without header row, assign column names manually
df = pd.read_excel(path, engine='openpyxl', header=None, skiprows=1)

col_names = [
    '业务品牌名称',    # 0
    '目标类型name',    # 1
    '日',              # 2
    'KA广告投放计划id', # 3
    '商家ID',          # 4
    '省份',            # 5  (all NaN!)
    '城市',            # 6  (actual city names)
    '投放计划名称',     # 7
    'KA账号id',        # 8
    '广告消耗',        # 9  (CA)
    '注册会员人数',     # 10
    '会员注册成本',     # 11
    '广告ROI',         # 12
    '交易会员人数',     # 13
    '广告点击次数',     # 14
    '广告曝光次数',     # 15
    '会员注册率',       # 16
    '广告CVR',         # 17
    '广告订单数',       # 18
    '广告CPM',         # 19
    '广告CPC',         # 20
    '广告CTR',         # 21
    '广告订单人数',     # 22
    '广告订单原价交易额', # 23
    '广告订单实付交易额', # 24 (GMV)
    '广告原价金额',     # 25
]
df.columns = col_names

print(f"Shape: {df.shape}")
print(f"\n=== 列名和类型 ===")
for c in col_names:
    print(f"  {c:20s} | non-null: {df[c].notna().sum():>6d} | type: {df[c].dtype}")

# Check unique cities
print(f"\n=== 城市维度 ===")
cities = df['城市'].dropna().unique()
print(f"唯一城市数: {len(cities)}")
print(f"前20个: {sorted(cities)[:20]}")
print(f"最后10个: {sorted(cities)[-10:]}")

# Check target types
print(f"\n=== 目标类型 ===")
print(df['目标类型name'].value_counts())

# Check 省份 column
print(f"\n=== 省份列 ===")
print(f"非空值: {df['省份'].notna().sum()}")
print(f"唯一值: {df['省份'].dropna().unique()}")

# Basic stats
print(f"\n=== 核心指标统计 ===")
print(f"总消耗(CA): {df['广告消耗'].sum():.2f}")
print(f"总注册会员: {df['注册会员人数'].sum():.0f}")
print(f"总GMV(实付): {df['广告订单实付交易额'].sum():.2f}")

# Check if there's existing "首销会员数" anywhere in data
print(f"\n=== 检查是否有首销会员数字段 ===")
# No 27th column exists, so no 首销会员数 for now

# Top 20 cities by CA
print(f"\n=== Top 20 城市 by CA(消耗) ===")
city_ca = df.groupby('城市')['广告消耗'].sum().sort_values(ascending=False)
print(city_ca.head(20).to_string())

# City × target type aggregation
print(f"\n=== Top 20 城市×目标类型 by CA ===")
tt_city = df.groupby(['目标类型name', '城市']).agg({
    '广告消耗': 'sum',
    '注册会员人数': 'sum',
    '广告订单实付交易额': 'sum',
    '广告ROI': 'mean',
}).sort_values('广告消耗', ascending=False)
print(tt_city.head(20).to_string())

print(f"\n=== 城市级数据(总览) ===")
city_all = df.groupby('城市').agg({
    '广告消耗': 'sum',
    '注册会员人数': 'sum',
    '广告订单实付交易额': 'sum',
    '广告ROI': 'mean',
}).sort_values('广告消耗', ascending=False)
city_all['CA/注册'] = city_all['广告消耗'] / city_all['注册会员人数'].replace(0, float('nan'))
city_all['ROI加权'] = city_all['广告订单实付交易额'] / city_all['广告消耗'].replace(0, float('nan'))
print(city_all.head(30).to_string())

# Check if there's city → province mapping needed
# The note from 好哥 says "城市补充对应的省份"
print(f"\n=== 中国城市→省份映射建议 ===")
# Check some municipal-level cities (直辖市)
municipal_direct = {'北京', '上海', '天津', '重庆'}
for c in sorted(cities):
    if c in municipal_direct:
        print(f"  {c} → 直辖市")
