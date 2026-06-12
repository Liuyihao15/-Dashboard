#!/usr/bin/env python3
"""Read brand member data using pandas (openpyxl), aggregate by city"""
import pandas as pd
import numpy as np

path = '/Users/edy/Desktop/品牌会员数据.xlsx'

# Read data - using openpyxl directly since xlsx is 35MB
# Only read specific columns by names
df = pd.read_excel(path, engine='openpyxl')

print(f"Shape: {df.shape}")
print(f"\nColumns ({len(df.columns)}):")
for i, c in enumerate(df.columns):
    print(f"  [{i:2d}] {c}")

print(f"\nFirst 3 rows:")
print(df.head(3).to_string())

# Check dtypes and basic stats
print(f"\n=== Dtypes ===")
print(df.dtypes)

# Check column 5 (省份) - unique values
col5 = df.columns[5]
print(f"\n=== Col 5 '{col5}' unique values ===")
print(f"Total unique: {df[col5].nunique()}")
print(f"Sample (first 30): {sorted(df[col5].unique())[:30]}")

# Check column 1 (目标类型) - unique values
col1 = df.columns[1]
print(f"\n=== Col 1 '{col1}' unique values ===")
print(df[col1].value_counts().to_string())

# Check if there's a 26th column (首销会员数)
if len(df.columns) > 25:
    print(f"\n=== Col 25 '{df.columns[25]}' preview ===")
    print(df.iloc[:, 25].describe())
else:
    print(f"\nNo 26th column found (total columns: {len(df.columns)})")

# Check col 9 (广告消耗/CA) and col 10 (注册会员数)
col9 = df.columns[9]
col10 = df.columns[10]
col11 = df.columns[11]
col24 = df.columns[24]

print(f"\n=== Col 9 '{col9}' (CA/消耗) ===")
print(f"Min: {df[col9].min()}, Max: {df[col9].max()}, Sum: {df[col9].sum():.2f}")
print(f"Non-zero rows: {(df[col9] > 0).sum()} / {len(df)}")
print(f"Median: {df[col9].median()}")
print(f"Top 10 values: {df[col9].nlargest(10).tolist()}")

print(f"\n=== Col 10 '{col10}' (注册会员数) ===")
print(f"Min: {df[col10].min()}, Max: {df[col10].max()}, Sum: {df[col10].sum()}")
print(f"Value counts (top 10):")
print(df[col10].value_counts().head(10).to_string())

print(f"\n=== Col 11 '{col11}' (会员注册成本) ===")
print(f"Min: {df[col11].min()}, Max: {df[col11].max()}, Mean: {df[col11].mean():.2f}")
print(f"Top 10: {df[col11].nlargest(10).tolist()}")

print(f"\n=== Col 24 '{col24}' (实付GMV) ===")
print(f"Min: {df[col24].min()}, Max: {df[col24].max()}, Sum: {df[col24].sum():.2f}")
print(f"Non-zero rows: {(df[col24] > 0).sum()} / {len(df)}")

# Aggregate by city
print(f"\n\n=== 按城市聚合 ===")
city_agg = df.groupby(col5).agg({
    col9: 'sum',        # 广告消耗 = CA
    col10: 'sum',       # 注册会员数
    col24: 'sum',       # 实付GMV
    col11: 'mean',      # 会员注册成本
    df.columns[12]: 'mean',  # ROI
}).sort_values(col9, ascending=False)

city_agg.columns = ['消耗(CA)', '注册会员', '实付GMV', '平均注册成本', '平均ROI']
print(f"\nTop 20 cities by CA:")
print(city_agg.head(20).to_string())

# Also aggregate by target type × city
print(f"\n\n=== 按目标类型 × 城市聚合 ===")
tt_city = df.groupby([col1, col5]).agg({
    col9: 'sum',
    col10: 'sum',
    col24: 'sum',
    df.columns[12]: 'mean',
}).sort_values(col9, ascending=False)

tt_city.columns = ['消耗(CA)', '注册会员', '实付GMV', '平均ROI']
print(f"\nTop 20 (type, city) by CA:")
print(tt_city.head(20).to_string())
