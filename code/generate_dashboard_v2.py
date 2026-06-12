#!/usr/bin/env python3
"""
品牌会员Dashboard - V2 (修复字体和emoji)
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import shutil

OUTPUT_DIR = '/Users/edy/Desktop/品牌会员Dashboard'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 尝试多种字体路径
FONT_CANDIDATES = [
    '/System/Library/Fonts/PingFang.ttc',
    '/System/Library/Fonts/STHeiti Light.ttc',
    '/System/Library/Fonts/Supplemental/Songti.ttc',
    '/System/Library/Fonts/Supplemental/STHeiti.ttf',
    '/System/Library/Fonts/Hiragino Sans GB.ttc',
    '/System/Library/Fonts/Supplemental/Arial Unicode.ttf',
]

for fp in FONT_CANDIDATES:
    if os.path.exists(fp):
        import matplotlib.font_manager as fm
        fm.fontManager.addfont(fp)
        print(f"Found font: {fp}")

# Try different font family names
for fname in ['PingFang SC', 'Heiti SC', 'STHeiti', 'Songti SC', 'Hiragino Sans GB', 'Arial Unicode MS']:
    try:
        plt.rcParams['font.family'] = fname
        fig, ax = plt.subplots(figsize=(2, 1))
        t = ax.text(0.5, 0.5, '测试Test', fontsize=12, ha='center')
        fig.canvas.draw()
        rendered = t.get_window_extent()
        plt.close()
        if rendered.width > 10:
            print(f"Font OK: {fname}")
            break
    except:
        plt.close()
        continue

plt.rcParams['axes.unicode_minus'] = False

# ====== 数据读取 ======
FILEPATH = '/Users/edy/Desktop/品牌会员数据.xlsx'

CITY_TO_PROVINCE = {
    '北京': '北京', '上海': '上海', '天津': '天津', '重庆': '重庆',
    '广州': '广东', '深圳': '广东', '东莞': '广东', '佛山': '广东',
    '珠海': '广东', '中山': '广东', '惠州': '广东', '汕头': '广东',
    '江门': '广东', '湛江': '广东', '肇庆': '广东', '茂名': '广东',
    '梅州': '广东', '汕尾': '广东', '河源': '广东', '阳江': '广东',
    '清远': '广东', '潮州': '广东', '揭阳': '广东', '云浮': '广东',
    '韶关': '广东',
    '南京': '江苏', '苏州': '江苏', '无锡': '江苏', '常州': '江苏',
    '南通': '江苏', '徐州': '江苏', '扬州': '江苏', '盐城': '江苏',
    '镇江': '江苏', '泰州': '江苏', '淮安': '江苏', '连云港': '江苏',
    '宿迁': '江苏',
    '杭州': '浙江', '宁波': '浙江', '温州': '浙江', '绍兴': '浙江',
    '嘉兴': '浙江', '金华': '浙江', '台州': '浙江', '湖州': '浙江',
    '丽水': '浙江', '衢州': '浙江', '舟山': '浙江',
    '成都': '四川', '绵阳': '四川', '德阳': '四川', '宜宾': '四川',
    '南充': '四川', '泸州': '四川', '眉山': '四川', '乐山': '四川',
    '自贡': '四川', '达州': '四川', '广元': '四川', '遂宁': '四川',
    '内江': '四川', '广安': '四川', '资阳': '四川', '雅安': '四川',
    '巴中': '四川', '攀枝花': '四川',
    '凉山彝族自治州': '四川', '甘孜藏族自治州': '四川',
    '福州': '福建', '厦门': '福建', '泉州': '福建', '漳州': '福建',
    '莆田': '福建', '宁德': '福建', '三明': '福建', '南平': '福建',
    '龙岩': '福建',
    '济南': '山东', '青岛': '山东', '烟台': '山东', '潍坊': '山东',
    '临沂': '山东', '淄博': '山东', '济宁': '山东', '德州': '山东',
    '泰安': '山东', '菏泽': '山东', '滨州': '山东', '威海': '山东',
    '东营': '山东', '枣庄': '山东', '日照': '山东', '聊城': '山东',
    '郑州': '河南', '洛阳': '河南', '南阳': '河南', '新乡': '河南',
    '商丘': '河南', '周口': '河南', '信阳': '河南', '安阳': '河南',
    '平顶山': '河南', '开封': '河南', '焦作': '河南', '许昌': '河南',
    '漯河': '河南', '三门峡': '河南', '濮阳': '河南', '鹤壁': '河南',
    '驻马店': '河南', '济源市': '河南',
    '武汉': '湖北', '襄阳': '湖北', '宜昌': '湖北', '荆州': '湖北',
    '黄冈': '湖北', '孝感': '湖北', '十堰': '湖北', '咸宁': '湖北',
    '恩施土家族苗族自治州': '湖北', '黄石': '湖北', '荆门': '湖北',
    '随州': '湖北', '潜江': '湖北', '仙桃': '湖北',
    '长沙': '湖南', '衡阳': '湖南', '株洲': '湖南', '岳阳': '湖南',
    '常德': '湖南', '湘潭': '湖南', '郴州': '湖南',
    '湘西土家族苗族自治州': '湖南',
    '合肥': '安徽', '芜湖': '安徽', '安庆': '安徽', '滁州': '安徽',
    '阜阳': '安徽', '马鞍山': '安徽', '蚌埠': '安徽', '宿州': '安徽',
    '六安': '安徽', '淮南': '安徽', '宣城': '安徽', '亳州': '安徽',
    '黄山': '安徽', '淮北': '安徽', '铜陵': '安徽', '池州': '安徽',
    '石家庄': '河北', '保定': '河北', '唐山': '河北', '邯郸': '河北',
    '沧州': '河北', '廊坊': '河北', '邢台': '河北', '秦皇岛': '河北',
    '衡水': '河北', '张家口': '河北', '承德': '河北',
    '西安': '陕西', '咸阳': '陕西', '宝鸡': '陕西', '渭南': '陕西',
    '延安': '陕西', '榆林': '陕西', '汉中': '陕西', '安康': '陕西',
    '商洛': '陕西', '铜川': '陕西',
    '南宁': '广西', '桂林': '广西', '柳州': '广西', '梧州': '广西',
    '北海': '广西', '玉林': '广西', '贵港': '广西', '百色': '广西',
    '钦州': '广西', '河池': '广西', '防城港': '广西', '崇左': '广西',
    '贺州': '广西', '来宾': '广西',
    '南昌': '江西', '赣州': '江西', '九江': '江西', '上饶': '江西',
    '宜春': '江西', '吉安': '江西', '抚州': '江西', '景德镇': '江西',
    '萍乡': '江西', '鹰潭': '江西', '新余': '江西',
    '沈阳': '辽宁', '大连': '辽宁', '鞍山': '辽宁', '锦州': '辽宁',
    '抚顺': '辽宁', '丹东': '辽宁', '营口': '辽宁', '盘锦': '辽宁',
    '辽阳': '辽宁', '铁岭': '辽宁', '葫芦岛': '辽宁',
    '太原': '山西', '晋中': '山西', '大同': '山西', '运城': '山西',
    '临汾': '山西', '吕梁': '山西', '忻州': '山西', '长治市': '山西',
    '晋城': '山西', '朔州': '山西',
    '贵阳': '贵州', '遵义': '贵州', '毕节': '贵州', '六盘水': '贵州',
    '铜仁': '贵州', '黔东南苗族侗族自治州': '贵州',
    '黔南布依族苗族自治州': '贵州', '黔西南布依族苗族自治州': '贵州',
    '安顺': '贵州',
    '昆明': '云南', '大理白族自治州': '云南', '红河州': '云南',
    '曲靖': '云南', '玉溪': '云南', '文山州': '云南',
    '楚雄彝族自治州': '云南', '普洱': '云南',
    '西双版纳傣族自治州': '云南', '德宏傣族景颇族自治州': '云南',
    '丽江': '云南', '保山': '云南', '临沧': '云南', '昭通': '云南',
    '迪庆藏族自治州': '云南',
    '哈尔滨': '黑龙江', '齐齐哈尔': '黑龙江', '大庆': '黑龙江',
    '牡丹江': '黑龙江', '绥化': '黑龙江', '佳木斯': '黑龙江',
    '黑河': '黑龙江', '鸡西': '黑龙江', '鹤岗': '黑龙江',
    '伊春市': '黑龙江', '大兴安岭地区': '黑龙江',
    '长春': '吉林', '吉林': '吉林', '通化': '吉林',
    '延边朝鲜族自治州': '吉林', '四平': '吉林', '白城': '吉林',
    '辽源': '吉林', '松原': '吉林', '白山': '吉林',
    '兰州': '甘肃', '天水': '甘肃', '庆阳': '甘肃', '酒泉': '甘肃',
    '武威': '甘肃', '定西': '甘肃', '陇南': '甘肃', '白银': '甘肃',
    '张掖': '甘肃', '平凉': '甘肃', '金昌': '甘肃', '嘉峪关': '甘肃',
    '甘南藏族自治州': '甘肃',
    '呼和浩特': '内蒙古', '包头': '内蒙古', '鄂尔多斯': '内蒙古',
    '呼伦贝尔': '内蒙古', '赤峰': '内蒙古', '通辽': '内蒙古',
    '乌兰察布': '内蒙古', '巴彦淖尔': '内蒙古', '乌海': '内蒙古',
    '兴安盟': '内蒙古', '锡林郭勒盟': '内蒙古', '阿拉善盟': '内蒙古',
    '乌鲁木齐': '新疆', '喀什地区': '新疆', '伊犁哈萨克自治州': '新疆',
    '巴音郭楞蒙古自治州': '新疆', '昌吉回族自治州': '新疆',
    '阿克苏地区': '新疆', '哈密地区': '新疆', '吐鲁番': '新疆',
    '和田地区': '新疆', '塔城地区': '新疆', '阿勒泰地区': '新疆',
    '博尔塔拉蒙古自治州': '新疆', '克孜勒苏柯尔克孜自治州': '新疆',
    '克拉玛依': '新疆', '石河子': '新疆', '阿拉尔市': '新疆',
    '五家渠市': '新疆', '图木舒克市': '新疆',
    '海口': '海南', '三亚': '海南', '儋州市': '海南', '琼海市': '海南',
    '万宁': '海南', '文昌市': '海南', '五指山市': '海南', '东方市': '海南',
    '陵水黎族自治县': '海南', '乐东黎族自治县': '海南',
    '昌江黎族自治县': '海南', '保亭黎族苗族自治县': '海南',
    '白沙黎族自治县': '海南', '琼中黎族苗族自治县': '海南',
    '屯昌县': '海南', '定安县': '海南', '临高县': '海南',
    '西宁': '青海', '海东市': '青海', '海南藏族自治州': '青海',
    '海西蒙古族藏族自治州': '青海', '玉树藏族自治州': '青海',
    '银川': '宁夏', '吴忠': '宁夏', '石嘴山': '宁夏', '固原': '宁夏',
    '中卫': '宁夏',
    '拉萨': '西藏', '日喀则': '西藏', '昌都': '西藏', '林芝': '西藏',
    '山南地区': '西藏', '那曲地区': '西藏', '阿里地区': '西藏',
}

print("Reading data...")
df = pd.read_excel(FILEPATH, engine='openpyxl', header=None, skiprows=1)
col_names = ['品牌', '目标类型name', '日', '投放计划id', '商家ID', '省份NaN', '城市',
             '投放计划名', 'KA账号id', '广告消耗', '注册会员人数', '会员注册成本',
             '广告ROI', '交易会员人数', '广告点击次数', '广告曝光次数',
             '会员注册率', '广告CVR', '广告订单数', '广告CPM', '广告CPC',
             '广告CTR', '广告订单人数', '广告订单原价交易额',
             '广告订单实付交易额', '广告原价金额']
df.columns = col_names
df['省份'] = df['城市'].map(CITY_TO_PROVINCE)
df['日'] = pd.to_datetime(df['日'].astype(str), format='%Y%m%d')
df = df[df['省份'].notna()].copy()

# Colors
GREEN = '#2b9751'
GREEN_DARK = '#155349'
RED = '#d02b29'
BLUE = '#254d86'
BLUE_LIGHT = '#5b8dd9'
GRAY = '#8c8c8c'
BG = '#f6f8f9'

def save_chart(fig, name):
    path = os.path.join(OUTPUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    print(f"  OK: {name}")

# === Chart 1: KPI 卡片 ===
print("Chart 1: KPI Cards...")
total_ca = df['广告消耗'].sum()
total_members = df['注册会员人数'].sum()
total_gmv = df['广告订单实付交易额'].sum()
total_roi = total_gmv / total_ca if total_ca > 0 else 0
total_cac = total_ca / total_members if total_members > 0 else 0

fig, axes = plt.subplots(1, 4, figsize=(14, 3.2), facecolor='white')
fig.suptitle('品牌会员投放 Dashboard - KPI总览', fontsize=16, fontweight='bold', color=BLUE, y=1.02)
kpis = [
    ('总消耗(CA)', f'${total_ca/10000:.1f}万', f'{total_ca:,.0f}', GREEN),
    ('总注册会员', f'{total_members:,.0f}', f'日均{total_members/38:.0f}', BLUE),
    ('总实付GMV', f'${total_gmv/10000:.1f}万', f'ROI {total_roi:.2f}', GREEN_DARK),
    ('新增成本', f'CAC ${total_cac:.1f}', f'CA/人', RED if total_cac > 100 else GREEN),
]
for ax, (title, big, sub, color) in zip(axes, kpis):
    ax.set_facecolor('#f8f9fa')
    ax.text(0.5, 0.75, title, fontsize=11, ha='center', va='center', color='#666')
    ax.text(0.5, 0.45, big, fontsize=22, ha='center', va='center', fontweight='bold', color=color)
    ax.text(0.5, 0.15, sub, fontsize=9, ha='center', va='center', color='#999')
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis('off')
save_chart(fig, '01_kpi_cards.png')

# === Chart 2: 省份气泡图（拉新）===
print("Chart 2: Province Bubble...")
new_df = df[df['目标类型name'] == '最大化新会员注册'].copy()
prov_new = new_df.groupby('省份').agg({'广告消耗': 'sum', '注册会员人数': 'sum', '广告订单实付交易额': 'sum'}).reset_index()
prov_new['加权ROI'] = prov_new['广告订单实付交易额'] / prov_new['广告消耗'].replace(0, 1)
prov_new['CAC'] = (prov_new['广告消耗'] / prov_new['注册会员人数'].replace(0, 1)).round(1)

fig, ax = plt.subplots(figsize=(12, 8), facecolor='white')
ax.set_facecolor(BG)
plot_data = prov_new[prov_new['省份'] != '西藏'].copy()
plot_data['size'] = plot_data['加权ROI'] * plot_data['广告消耗'] / 500
scatter = ax.scatter(plot_data['广告消耗'], plot_data['注册会员人数'],
    s=plot_data['size'].clip(50, 2000), c=plot_data['CAC'],
    cmap='RdYlGn_r', alpha=0.7, edgecolors='white', linewidth=1, vmin=60, vmax=160)
for _, row in plot_data.iterrows():
    if row['省份'] in ['广东','江苏','北京','上海','山东','浙江','河南','河北','四川','重庆']:
        ax.annotate(row['省份'], (row['广告消耗'], row['注册会员人数']), fontsize=9, ha='center', va='bottom', fontweight='bold', color='#333')
    elif row['注册会员人数'] > 2000 or row['广告消耗'] > 150000:
        ax.annotate(row['省份'], (row['广告消耗'], row['注册会员人数']), fontsize=7, ha='center', va='bottom', color='#999')
median_ca, median_mem = plot_data['广告消耗'].median(), plot_data['注册会员人数'].median()
ax.axvline(median_ca, color='gray', ls='--', alpha=0.3, lw=0.8)
ax.axhline(median_mem, color='gray', ls='--', alpha=0.3, lw=0.8)
cbar = plt.colorbar(scatter, ax=ax, shrink=0.7); cbar.set_label('CAC ($/人)', fontsize=10)
ax.set_xlabel('广告消耗 CA ($)', fontsize=12, fontweight='bold')
ax.set_ylabel('注册会员人数', fontsize=12, fontweight='bold')
ax.set_title('省份拉新气泡图 - 最大化新会员注册\n气泡大小=加权ROI  颜色=CAC', fontsize=14, fontweight='bold', color=BLUE, pad=15)
save_chart(fig, '02_province_bubble.png')

# === Chart 3: GMV贡献省份排行 ===
print("Chart 3: GMV by Province...")
prov_all = df.groupby('省份').agg({'广告消耗': 'sum', '注册会员人数': 'sum', '广告订单实付交易额': 'sum'}).reset_index()
prov_all = prov_all.sort_values('广告订单实付交易额', ascending=True)
prov_all['ROI'] = (prov_all['广告订单实付交易额'] / prov_all['广告消耗'].replace(0, 1)).round(2)

fig, ax = plt.subplots(figsize=(12, 10), facecolor='white')
ax.set_facecolor(BG)
plot_df = prov_all.tail(31)
gmv_col = plot_df['广告订单实付交易额'] / 10000
bars = ax.barh(range(len(plot_df)), gmv_col, height=0.7)
for i, (_, row) in enumerate(plot_df.iterrows()):
    roi = row['ROI']
    bars[i].set_color(GREEN if roi >= 10 else (BLUE if roi >= 7.5 else RED))
    ax.text(gmv_col.iloc[i] + 3, i, f'ROI {roi:.1f}', va='center', fontsize=7, color='#555')
ax.set_yticks(range(len(plot_df)))
ax.set_yticklabels(plot_df['省份'], fontsize=9)
ax.set_xlabel('实付GMV (万元)', fontsize=11, fontweight='bold')
ax.set_title('GMV贡献 - 省份排行', fontsize=14, fontweight='bold', color=BLUE, pad=15)
ax.set_xlim(0, gmv_col.max() * 1.25)
ax.grid(axis='x', alpha=0.2)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
save_chart(fig, '03_gmv_by_province.png')

# === Chart 4: 每日趋势 ===
print("Chart 4: Daily Trend...")
daily = df.groupby('日').agg({'广告消耗': 'sum', '注册会员人数': 'sum'}).reset_index().sort_values('日')
fig, ax1 = plt.subplots(figsize=(14, 5), facecolor='white')
ax1.set_facecolor(BG)
ax1.fill_between(daily['日'], daily['广告消耗'], alpha=0.15, color=GREEN)
ax1.plot(daily['日'], daily['广告消耗'], color=GREEN, lw=2, marker='o', ms=3, label='CA(广告消耗)')
ax1.set_ylabel('广告消耗 CA ($)', fontsize=11, fontweight='bold', color=GREEN)
ax1.tick_params(axis='y', labelcolor=GREEN)
ax1.set_xlabel('日期', fontsize=11)
ax2 = ax1.twinx()
ax2.plot(daily['日'], daily['注册会员人数'], color=BLUE_LIGHT, lw=1.5, ls='--', marker='s', ms=2, label='注册会员数')
ax2.set_ylabel('注册会员数', fontsize=11, fontweight='bold', color=BLUE_LIGHT)
ax2.tick_params(axis='y', labelcolor=BLUE_LIGHT)
l1, la1 = ax1.get_legend_handles_labels(); l2, la2 = ax2.get_legend_handles_labels()
ax1.legend(l1+l2, la1+la2, loc='upper left', fontsize=10)
ax1.set_title('每日趋势 - CA消耗 vs 注册会员数', fontsize=14, fontweight='bold', color=BLUE, pad=15)
ax1.grid(alpha=0.2)
fig.autofmt_xdate()
save_chart(fig, '04_daily_trend.png')

# === Chart 5: ROI vs CAC 效率矩阵 ===
print("Chart 5: Efficiency Matrix...")
eff = df.groupby('省份').agg({'广告消耗': 'sum', '注册会员人数': 'sum', '广告订单实付交易额': 'sum'}).reset_index()
eff['ROI'] = (eff['广告订单实付交易额'] / eff['广告消耗'].replace(0, 1)).round(2)
eff['CAC'] = (eff['广告消耗'] / eff['注册会员人数'].replace(0, 1)).round(1)
eff = eff[eff['省份'] != '西藏']
fig, ax = plt.subplots(figsize=(12, 8), facecolor='white')
ax.set_facecolor(BG)
eff['size'] = eff['广告消耗'] / 5000
scatter = ax.scatter(eff['ROI'], eff['CAC'], s=eff['size'].clip(30, 1500),
    c=eff['广告消耗'], cmap='YlOrRd', alpha=0.7, edgecolors='white', linewidth=1)
for _, row in eff.iterrows():
    ax.annotate(row['省份'], (row['ROI'], row['CAC']), fontsize=8, ha='center', va='bottom', fontweight='bold')
ax.axvline(eff['ROI'].median(), color='gray', ls='--', alpha=0.3)
ax.axhline(eff['CAC'].median(), color='gray', ls='--', alpha=0.3)
cbar = plt.colorbar(scatter, ax=ax, shrink=0.7); cbar.set_label('CA消耗总额 $', fontsize=10)
ax.set_xlabel('加权ROI', fontsize=12, fontweight='bold')
ax.set_ylabel('CAC ($/注册会员)', fontsize=12, fontweight='bold')
ax.set_title('效率看板 - ROI x CAC矩阵\n气泡大小=消耗规模', fontsize=14, fontweight='bold', color=BLUE, pad=15)
ax.grid(alpha=0.2)
save_chart(fig, '05_efficiency_matrix.png')

# === Chart 6: 目标类型对比 ===
print("Chart 6: Target Comparison...")
tt = df.groupby('目标类型name').agg({'广告消耗': 'sum', '注册会员人数': 'sum', '广告订单实付交易额': 'sum'}).reset_index()
tt['ROI'] = (tt['广告订单实付交易额'] / tt['广告消耗'].replace(0, 1)).round(2)
tt['CAC'] = (tt['广告消耗'] / tt['注册会员人数'].replace(0, 1)).round(1)
labels = tt['目标类型name'].str.replace('最大化', '').str.replace('会员', '')
fig, axes = plt.subplots(1, 3, figsize=(14, 4.5), facecolor='white')
fig.suptitle('目标类型对比 - 拉新 vs 老会员转化', fontsize=14, fontweight='bold', color=BLUE, y=1.05)
colors = [GREEN, BLUE]
axes[0].pie(tt['广告消耗'], labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
axes[0].set_title('CA消耗占比', fontsize=11)
x = np.arange(2)
w = 0.3
axes[1].bar(x - w/2, tt['注册会员人数'], w, label='注册会员', color=GREEN)
axes[1].bar(x + w/2, tt['广告订单实付交易额']/100, w, label='GMV/100', color=BLUE_LIGHT)
axes[1].set_xticks(x); axes[1].set_xticklabels(labels, fontsize=8)
axes[1].set_title('注册会员 vs GMV', fontsize=11); axes[1].legend(fontsize=8)
roi_vals = tt['ROI'].values; cac_vals = tt['CAC'].values
axes[2].bar(x - w/2, roi_vals, w, label='ROI', color=GREEN)
axes[2].bar(x + w/2, cac_vals/10, w, label='CAC/10', color=BLUE)
axes[2].set_xticks(x); axes[2].set_xticklabels(labels, fontsize=8)
axes[2].set_title('ROI vs CAC', fontsize=11); axes[2].legend(fontsize=8)
for ax in axes: ax.grid(alpha=0.15); ax.set_facecolor(BG)
plt.tight_layout()
save_chart(fig, '06_target_comparison.png')

# === CSV 数据 ===
prov_csv = df.groupby('省份').agg({
    '广告消耗': 'sum', '注册会员人数': 'sum', '广告订单实付交易额': 'sum',
    '广告点击次数': 'sum', '广告曝光次数': 'sum'
}).reset_index().sort_values('广告消耗', ascending=False)
prov_csv['加权ROI'] = (prov_csv['广告订单实付交易额'] / prov_csv['广告消耗'].replace(0, 1)).round(2)
prov_csv['CAC($/人)'] = (prov_csv['广告消耗'] / prov_csv['注册会员人数'].replace(0, 1)).round(1)
prov_csv['CA占比(%)'] = (prov_csv['广告消耗'] / prov_csv['广告消耗'].sum() * 100).round(1)
prov_csv.to_csv(os.path.join(OUTPUT_DIR, 'province_data.csv'), index=False, encoding='utf-8-sig')

print(f"\nAll files in {OUTPUT_DIR}:")
for f in sorted(os.listdir(OUTPUT_DIR)):
    sz = os.path.getsize(os.path.join(OUTPUT_DIR, f))
    print(f"  {f} ({sz/1024:.1f}KB)")
