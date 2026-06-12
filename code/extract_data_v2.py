#!/usr/bin/env python3
"""重新提取：城市名列（col5）就是省份——直接映射到省份"""
import zipfile, re, json
from xml.etree import ElementTree as ET

# 已知是省份名的城市（col5中北京/上海/重庆/天津/吉林实际上是城市，需要映射）
# 其他如"广东"、"江苏"等本身就是省份名，直接使用
PROVINCE_NAMES = {'广东','江苏','浙江','河南','山东','河北','江西','四川',
    '湖北','湖南','安徽','福建','广西','陕西','云南','贵州','山西',
    '辽宁','黑龙江','甘肃','新疆','内蒙古','海南','青海','宁夏','西藏'}

# 直辖市/特别行政区 - 直接就是省份
MUNICIPAL = {'北京':'北京','上海':'上海','天津':'天津','重庆':'重庆'}

# 城市->省份的完整映射（从v2）
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
    '阿坝藏族羌族自治州': '四川',
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
    '常德': '湖南', '湘潭': '湖南', '郴州': '湖南', '邵阳': '湖南',
    '永州': '湖南', '怀化': '湖南', '娄底': '湖南', '益阳': '湖南',
    '张家界': '湖南',
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
    '迪庆藏族自治州': '云南', '怒江傈僳族自治州': '云南',
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

def resolve_province(city_value):
    """解析省份：city_value可能是省份名、城市名或直辖市"""
    if not city_value or city_value in ('0','NULL','','None','nan','#N/A'):
        return None
    # 如果本身就是省份名，直接返回
    if city_value in PROVINCE_NAMES:
        return city_value
    # 如果是直辖市
    if city_value in MUNICIPAL:
        return MUNICIPAL[city_value]
    # 通过城市映射
    if city_value in CITY_TO_PROVINCE:
        return CITY_TO_PROVINCE[city_value]
    # 特处理带括号的
    for prefix in ['内蒙古', '黑龙江', '新疆', '广西']:
        if city_value.startswith(prefix):
            return prefix
    return None


with zipfile.ZipFile('/Users/edy/Desktop/品牌会员数据.xlsx', 'r') as z:
    tree = ET.parse(z.open('xl/sharedStrings.xml'))
    root = tree.getroot()
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sis = root.findall('.//s:si', ns)
    strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in sis]
    
    raw = z.open('xl/worksheets/sheet1.xml').read().decode()
    rows = re.findall(r'<row[^>]*>(.*?)</row>', raw, re.DOTALL)

# Columns (0-based):
# 0:品牌 1:目标类型name 2:日 3:投放计划id 4:商家ID 5:城市/省份 6:城市编码
# 9:广告消耗 10:注册会员人数 11:会员注册成本 12:广告ROI
# 18:广告订单数 23:广告订单实付交易额

daily = {}  # date -> {acq:{cost,members,gmv,orders}, ret:{cost,members}}
city_stats = {}
prov_stats = {}
prov_new_stats = {}
prov_ret_stats = {}
unmapped = set()
all_province_names = set()

for row_xml in rows[1:]:
    cells = re.findall(r'(<c[^>]*>.*?</c>)', row_xml, re.DOTALL)
    row = {}
    for cell_str in cells:
        rm = re.search(r'r="([^"]+)"', cell_str)
        if not rm: continue
        ref = rm.group(1)
        m2 = re.match(r'([A-Z]+)', ref)
        if not m2: continue
        col = 0
        for c in m2.group(1):
            col = col * 26 + (ord(c) - ord('A') + 1)
        col -= 1
        is_str = 't="s"' in cell_str
        vm = re.search(r'<v>(.*?)</v>', cell_str)
        if vm:
            val = vm.group(1)
            if is_str:
                try: val = strings[int(val)]
                except: pass
            if val == 'NULL':
                row[col] = 0
            else:
                try: row[col] = float(val) if '.' in str(val) or not is_str else val
                except: row[col] = val
    
    target = str(row.get(1, '')).strip()
    date_str = str(row.get(2, '')).strip()
    city_val = str(row.get(5, '')).strip()
    
    cost = float(row.get(9, 0) or 0)
    members = float(row.get(10, 0) or 0)
    gmv = float(row.get(23, 0) or 0)
    orders = float(row.get(18, 0) or 0)
    
    # Province resolution
    prov = resolve_province(city_val)
    if not prov:
        unmapped.add(city_val)
        continue
    
    all_province_names.add(prov)
    
    # Province-level stats
    if prov not in prov_stats:
        prov_stats[prov] = {'cost': 0, 'members': 0, 'gmv': 0}
        prov_new_stats[prov] = {'cost': 0, 'members': 0, 'gmv': 0}
        prov_ret_stats[prov] = {'cost': 0, 'members': 0}
    prov_stats[prov]['cost'] += cost
    prov_stats[prov]['members'] += members
    prov_stats[prov]['gmv'] += gmv
    if '最大化新会员注册' in target:
        prov_new_stats[prov]['cost'] += cost
        prov_new_stats[prov]['members'] += members
        prov_new_stats[prov]['gmv'] += gmv
    else:
        prov_ret_stats[prov]['cost'] += cost
        prov_ret_stats[prov]['members'] += members
    
    # City-level stats (use city_val as city key)
    if city_val and city_val not in ('0','NULL','','None','nan','#N/A'):
        if city_val not in city_stats:
            city_stats[city_val] = {'cost': 0, 'members': 0, 'gmv': 0, 'roi_sum': 0, 'count': 0, 'prov': prov}
        city_stats[city_val]['cost'] += cost
        city_stats[city_val]['members'] += members
        city_stats[city_val]['gmv'] += gmv
        if cost > 0:
            city_stats[city_val]['roi_sum'] += gmv / cost
            city_stats[city_val]['count'] += 1
    
    # Daily by target type
    if date_str and len(date_str) >= 8:
        date_key = date_str[:8]
        if date_key not in daily:
            daily[date_key] = {'acq': {'cost':0,'members':0,'gmv':0,'orders':0}, 'ret': {'cost':0,'members':0}}
        if '最大化新会员注册' in target:
            daily[date_key]['acq']['cost'] += cost
            daily[date_key]['acq']['members'] += members
            daily[date_key]['acq']['gmv'] += gmv
            daily[date_key]['acq']['orders'] += orders
        else:
            daily[date_key]['ret']['cost'] += cost
            daily[date_key]['ret']['members'] += members

if unmapped:
    print(f"Unmapped ({len(unmapped)}): {sorted(unmapped)[:30]}")

dates_sorted = sorted(daily.keys())

# Build output
out = {
    'dates': dates_sorted,
    'acq_cost': [round(daily[d]['acq']['cost'], 2) for d in dates_sorted],
    'acq_members': [int(daily[d]['acq']['members']) for d in dates_sorted],
    'acq_gmv': [round(daily[d]['acq']['gmv'], 2) for d in dates_sorted],
    'ret_cost': [round(daily[d]['ret']['cost'], 2) for d in dates_sorted],
    'ret_members': [int(daily[d]['ret']['members']) for d in dates_sorted],
}

# Province data (sorted by cost descending)
prov_names = sorted(prov_stats.keys(), key=lambda p: prov_stats[p]['cost'], reverse=True)
out['provinces'] = prov_names
out['prov_cost'] = [round(prov_stats[p]['cost'], 2) for p in prov_names]
out['prov_members'] = [int(prov_stats[p]['members']) for p in prov_names]
out['prov_gmv'] = [round(prov_stats[p]['gmv'], 2) for p in prov_names]
out['prov_cac'] = [round(prov_stats[p]['cost']/prov_stats[p]['members'], 2) if prov_stats[p]['members'] > 0 else 0 for p in prov_names]
out['prov_roi'] = [round(prov_stats[p]['gmv']/prov_stats[p]['cost'], 2) if prov_stats[p]['cost'] > 0 else 0 for p in prov_names]
out['prov_new_cost'] = [round(prov_new_stats[p]['cost'], 2) for p in prov_names]
out['prov_new_members'] = [int(prov_new_stats[p]['members']) for p in prov_names]
out['prov_new_gmv'] = [round(prov_new_stats[p]['gmv'], 2) for p in prov_names]
out['prov_ret_cost'] = [round(prov_ret_stats[p]['cost'], 2) for p in prov_names]
out['prov_ret_members'] = [int(prov_ret_stats[p]['members']) for p in prov_names]

# City data (sorted by cost descending)
city_names = sorted(city_stats.keys(), key=lambda c: city_stats[c]['cost'], reverse=True)
out['cities'] = city_names
out['city_cost'] = [round(city_stats[c]['cost'], 2) for c in city_names]
out['city_members'] = [int(city_stats[c]['members']) for c in city_names]
out['city_gmv'] = [round(city_stats[c]['gmv'], 2) for c in city_names]
out['city_roi'] = [round(city_stats[c]['gmv']/city_stats[c]['cost'], 2) if city_stats[c]['cost'] > 0 else 0 for c in city_names]
out['city_cac'] = [round(city_stats[c]['cost']/city_stats[c]['members'], 2) if city_stats[c]['members'] > 0 else 0 for c in city_names]
out['city_prov'] = [city_stats[c]['prov'] for c in city_names]

# KPIs
total_cost = sum(out['prov_cost'])
total_members = sum(out['prov_members'])
total_gmv = sum(out['prov_gmv'])
out['kpi'] = {
    'total_cost': round(total_cost, 2),
    'total_members': total_members,
    'total_gmv': round(total_gmv, 2),
    'total_roi': round(total_gmv / total_cost, 2) if total_cost > 0 else 0,
    'total_cac': round(total_cost / total_members, 2) if total_members > 0 else 0,
    'province_count': len(prov_names),
    'city_count': len(city_names),
}

# Acq/Ret totals
acq_total_cost = sum(out['acq_cost'])
acq_total_members = sum(out['acq_members'])
ret_total_cost = sum(out['ret_cost'])
ret_total_members = sum(out['ret_members'])
out['kpi']['acq_cost'] = round(acq_total_cost, 2)
out['kpi']['acq_members'] = acq_total_members
out['kpi']['acq_cac'] = round(acq_total_cost / acq_total_members, 2) if acq_total_members > 0 else 0
out['kpi']['ret_cost'] = round(ret_total_cost, 2)
out['kpi']['ret_members'] = ret_total_members
out['kpi']['ret_cac'] = round(ret_total_cost / ret_total_members, 2) if ret_total_members > 0 else 0

print(json.dumps(out, ensure_ascii=False))
