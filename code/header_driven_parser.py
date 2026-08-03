#!/usr/bin/env python3
"""
霸王茶姬看板 表头驱动解析器（2026-08-03 新增，替代旧版固定列下标 update_dashboard.py）

用途：正确解析"品牌会员-霸王_YYYYMMDD-YYYYMMDD(日).xlsx"。
关键：不同周导出文件的列序不同（如上周文件 广告消耗 在 col9、本周在 col8），
     必须按列名动态定位，不能硬编码列下标，否则数据错位。

用法：
  python3 header_driven_parser.py <xlsx路径>
  输出 daily 每日汇总(acq/ret分列+org/real GMV)

兼容性说明：
  - 主明细自动选"行数最多的 sheet"（上周在sheet2、本周在sheet1）
  - 自动处理 NULL / 空值
  - 自动按 目标类型name 含"新会员" 区分 acq / ret
"""
import zipfile, xml.etree.ElementTree as ET, sys, re


def colidx(cs):
    c = 0
    for ch in ''.join([x for x in cs if not x.isdigit()]):
        c = c * 26 + (ord(ch) - ord('A') + 1)
    return c - 1


def clean(name):
    return re.sub(r'[\s（）()]', '', name or '')


def parse_week(xlsx_path):
    """表头驱动解析每日数据，返回 daily[date] = {acq_cost,acq_mem,acq_trade,acq_order,ret_cost,ret_mem,ret_trade,ret_order,org_gmv,real_gmv}"""
    z = zipfile.ZipFile(xlsx_path)
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    try:
        st = ET.fromstring(z.read('xl/sharedStrings.xml'))
        strings = [si.find('.//s:t', ns).text if si.find('.//s:t', ns) is not None else '' for si in st.findall('.//s:si', ns)]
    except Exception:
        strings = {}
    # 主明细 = 行数最多的 sheet
    best_sheet, best_rows = None, 0
    for n in z.namelist():
        if n.startswith('xl/worksheets/') and n.endswith('.xml'):
            root = ET.fromstring(z.read(n))
            nr = len(root.findall('.//s:row', ns))
            if nr > best_rows:
                best_rows, best_sheet = nr, n
    root = ET.fromstring(z.read(best_sheet))
    rows = root.findall('.//s:row', ns)

    # 表头→列号 动态映射
    header = {}
    for cc in rows[0].findall('s:c', ns):
        ci = colidx(cc.get('r', ''))
        typ = cc.get('t', '')
        v = cc.find('s:v', ns)
        if v is not None:
            header[clean(strings[int(v.text)] if typ == 's' else v.text)] = ci

    def H(*names):
        for nn in names:
            if clean(nn) in header:
                return header[clean(nn)]
        return None

    c_date = H('日')
    c_tgt = H('目标类型name', '目标类型')
    c_cost = H('广告消耗')
    c_mem = H('注册会员人数（直接）', '注册会员人数')
    c_trade = H('交易会员人数')
    c_ord = H('广告订单人数')
    c_org = H('广告订单原价交易额')
    c_real = H('广告订单实付交易额')
    if c_cost is None:
        print('❌ 未找到"广告消耗"列，表头:', list(header.keys()))
        return {}

    def sf(v):
        try:
            return float(v) if str(v).strip().upper() not in ('NULL', 'NONE', 'NAN', '', 'INF') else 0.0
        except (ValueError, TypeError):
            return 0.0

    daily = {}
    for r in rows[1:]:
        vals = {}
        for cc in r.findall('s:c', ns):
            ci = colidx(cc.get('r', ''))
            typ = cc.get('t', '')
            v = cc.find('s:v', ns)
            if v is not None:
                vals[ci] = strings[int(v.text)] if typ == 's' else v.text
        d = str(vals.get(c_date, '') or '').strip()
        if not d or not d.isdigit():
            continue
        is_acq = '新会员' in str(vals.get(c_tgt, ''))
        if d not in daily:
            daily[d] = {'acq_cost': 0, 'acq_mem': 0, 'acq_trade': 0, 'acq_order': 0,
                        'ret_cost': 0, 'ret_mem': 0, 'ret_trade': 0, 'ret_order': 0,
                        'org_gmv': 0, 'real_gmv': 0}
        cost, mem = sf(vals.get(c_cost)), sf(vals.get(c_mem))
        trade, ords = sf(vals.get(c_trade)), sf(vals.get(c_ord))
        org, real = sf(vals.get(c_org)), sf(vals.get(c_real))
        daily[d]['org_gmv'] += org
        daily[d]['real_gmv'] += real
        if is_acq:
            daily[d]['acq_cost'] += cost
            daily[d]['acq_mem'] += mem
            daily[d]['acq_trade'] += trade
            daily[d]['acq_order'] += ords
        else:
            daily[d]['ret_cost'] += cost
            daily[d]['ret_mem'] += mem
            daily[d]['ret_trade'] += trade
            daily[d]['ret_order'] += ords
    return daily


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python3 header_driven_parser.py <xlsx路径>')
        sys.exit(1)
    daily = parse_week(sys.argv[1])
    print(f'解析到 {len(daily)} 天')
    for d in sorted(daily):
        v = daily[d]
        print(f'  {d}: acq_cost={v["acq_cost"]:,.0f} acq_mem={v["acq_mem"]:,.0f} '
              f'ret_cost={v["ret_cost"]:,.0f} org_gmv={v["org_gmv"]:,.0f} '
              f'acq_order={v["acq_order"]:,.0f} acq_trade={v["acq_trade"]:,.0f}')
