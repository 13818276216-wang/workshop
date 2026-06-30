# -*- coding: utf-8 -*-
import sys, json, math
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from collections import defaultdict
from datetime import datetime

# ========== 读数据 ==========
# 数据源：飞书在线表格 https://daixiansheng.feishu.cn/file/Ju6Nb7V7LoEfmPxVVCEcR5WFnZf
# 先通过 lark-cli drive +download --file-token "Ju6Nb7V7LoEfmPxVVCEcR5WFnZf" --output "sales_detail.feishu.xlsx" --as user 下载
path = r"E:\袋鼠先生\sales_detail.feishu.xlsx"
wb = openpyxl.load_workbook(path, data_only=True)
ws = wb.active

headers = [cell.value for cell in next(ws.iter_rows(max_row=1))]
col = {h: i for i, h in enumerate(headers)}

rows_data = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        continue
    rows_data.append(row)
wb.close()

print(f"总数据行: {len(rows_data)}")

# ========== 剔除鸣鸣很忙 ==========
def is_mingmang(name):
    """判断是否鸣鸣很忙体系客户"""
    if not name:
        return False
    keywords = ['零食很忙', '赵一鸣', '鸣忙', '很忙零食', '长沙晓忙']
    for kw in keywords:
        if kw in str(name):
            return True
    return False

filtered = [r for r in rows_data if not is_mingmang(r[col['客户名称']])]
print(f"剔除鸣鸣很忙后: {len(filtered)} 行")

# ========== 读合同台账获取省区经理 ==========
# 数据源：飞书在线表格「线下记录（授权书/证明关系）」→ sheet「合同台账」
# URL: https://daishuxiansheng.feishu.cn/wiki/PZgMwdwlGimD60kypFOcBVasn9c?sheet=M4qpLc
import subprocess, tempfile, csv, io
LARK_CLI = r'C:\Users\Administrator\.workbuddy\binaries\node\cli-connector-packages\lark-cli.cmd'
try:
    result = subprocess.run([
        LARK_CLI, 'sheets', '+csv-get',
        '--url', 'https://daishuxiansheng.feishu.cn/wiki/PZgMwdwlGimD60kypFOcBVasn9c',
        '--sheet-id', 'M4qpLc',
        '--range', 'A1:N61'
    ], capture_output=True, text=True, encoding='utf-8')
    import json as _json
    resp = _json.loads(result.stdout)
    csv_text = resp['data']['annotated_csv']
    # 按 [row=N] 前缀切分行，处理含换行的单元格
    import re
    raw_lines = csv_text.strip().split('\n')
    rows = []
    current = None
    for line in raw_lines:
        m = re.match(r'\[row=(\d+)\]\s(.*)', line)
        if m:
            if current is not None:
                rows.append(current)
            current = m.group(2)
        else:
            if current is not None:
                current += '\n' + line
    if current is not None:
        rows.append(current)
    # 用csv解析每行
    reader = csv.reader(io.StringIO('\n'.join(rows)))
    contract_headers = next(reader)
    contract_col = {h: i for i, h in enumerate(contract_headers)}
    
    manager_map = {} # 客户名称 -> 负责人
    for row in reader:
        if len(row) < 12:
            continue
        cname = str(row[contract_col.get('合作方', 2)] or '').strip()
        manager = str(row[contract_col.get('负责人', 11)] or '').strip()
        if cname and manager and manager != 'None':
            manager_map[cname] = manager
    print(f"成功从飞书读取合同台账，匹配到 {len(manager_map)} 个负责人的映射。")
except Exception as e:
    print("读取飞书合同台账失败:", e)
    manager_map = {}

# ========== 解析数据 ==========
from datetime import datetime as dt
now = dt.now()
current_month_calc = now.strftime('%Y-%m')

# 先扫一遍所有时间戳，找出数据里的最后一天作为 today_key_calc
channel_idx = col['销售渠道']  # 经销商维度：销售渠道
cname_idx = col['客户名称']     # 仅用于剔除鸣鸣很忙
pname_idx = col['货品名称']
qty_idx = col['数量']
price_idx = col['单价']
amt_idx = col['金额']
profit_idx = col['毛利']
time_idx = col['货品级发货时间']
order_idx = col['订单编号']

all_times = []
for r in filtered:
    ts = str(r[time_idx] or '').strip()
    if ts:
        all_times.append(ts[:10])
all_days = sorted(set(all_times))
today_key_calc = all_days[-1] if all_days else ''

print(f"数据最后日期: {today_key_calc}, 当月: {current_month_calc}")

dealers = {}       # dealer_name -> {销售额, 毛利, 订单数, 负责人}
dealers_today = defaultdict(float)  # dealer_name -> 今日销售额
dealers_month = defaultdict(float)  # dealer_name -> 本月销售额
dealers_first_month = {}  # dealer_name -> 首批订单月(YYYY-MM)
managers = {}      # manager_name -> {销售额, 毛利, 订单数, 负责的客户数, 今日销售额, 本月销售额, 本月毛利}
managers_today = defaultdict(float)  # manager_name -> 今日销售额
managers_month = defaultdict(lambda: {'sales': 0, 'profit': 0})  # manager_name -> 本月销售额/毛利
products = {}      # product_name -> {销售额, 销量, 毛利, 单价列表}
monthly = defaultdict(lambda: {'sales': 0, 'profit': 0, 'orders': set()})
daily = defaultdict(lambda: {'sales': 0, 'orders': set()}) # 增加日维度
total_sales = 0
total_profit = 0
total_orders = set()
total_dealers = set()

cname_idx = col['客户名称']
channel_idx = col['销售渠道']
pname_idx = col['货品名称']
qty_idx = col['数量']
price_idx = col['单价']
amt_idx = col['金额']
profit_idx = col['毛利']
time_idx = col['货品级发货时间']
order_idx = col['订单编号']

for r in filtered:
    channel = str(r[channel_idx] or '').strip()  # 经销商维度：销售渠道
    cname = str(r[cname_idx] or '').strip()       # 客户名称，用于经理映射
    pname = str(r[pname_idx] or '').strip()
    qty = float(r[qty_idx] or 0)
    price = float(r[price_idx] or 0)
    amt = float(r[amt_idx] or 0)
    profit = float(r[profit_idx] or 0) if r[profit_idx] is not None else 0
    order_id = str(r[order_idx] or '')
    time_str = str(r[time_idx] or '')

    if qty == 0 and amt == 0:
        continue

    pos_qty = max(qty, 0)
    pos_amt = max(amt, 0)
    # 修正异常客户名称：黄辉15257127683 按销售渠道映射回正常客户名称
    cname_fix_map = {
        '杭州品强电子商务有限公司': '杭州品强电子商务有限公司',
        '西安海盒鲜食品有限公司': '西安海盒鲜食品有限公司',
        '郑州锦门商贸有限公司': '郑州锦门商贸有限公司',
    }
    if cname == '黄辉15257127683':
        cname = cname_fix_map.get(channel, cname)
    manager = manager_map.get(cname, '未知')
    
    if channel not in dealers:
        dealers[channel] = {'sales': 0, 'profit': 0, 'orders': set(), 'qty': 0, 'manager': manager}
    dealers[channel]['sales'] += pos_amt
    dealers[channel]['profit'] += profit
    dealers[channel]['orders'].add(order_id)
    dealers[channel]['qty'] += pos_qty

    if manager not in managers:
        managers[manager] = {'sales': 0, 'profit': 0, 'orders': set(), 'clients': set()}
    managers[manager]['sales'] += pos_amt
    managers[manager]['profit'] += profit
    managers[manager]['orders'].add(order_id)
    managers[manager]['clients'].add(channel)

    if pname not in products:
        products[pname] = {'sales': 0, 'qty': 0, 'profit': 0, 'prices': [], 'dealers': set()}
    products[pname]['sales'] += pos_amt
    products[pname]['qty'] += pos_qty
    products[pname]['profit'] += profit
    if price > 0:
        products[pname]['prices'].append(price)
    products[pname]['dealers'].add(channel)

    if time_str:
        try:
            month_key = time_str[:7]
            monthly[month_key]['sales'] += pos_amt
            monthly[month_key]['profit'] += profit
            monthly[month_key]['orders'].add(order_id)
            
            day_key = time_str[:10] # YYYY-MM-DD
            daily[day_key]['sales'] += pos_amt
            daily[day_key]['orders'].add(order_id)
            
            # 按省区经理维度：今日 + 本月
            if day_key == today_key_calc:
                managers_today[manager] += pos_amt
            if month_key == current_month_calc:
                managers_month[manager]['sales'] += pos_amt
                managers_month[manager]['profit'] += profit
            
            # 按经销商维度：今日 + 本月 + 首批月
            if day_key == today_key_calc:
                dealers_today[channel] += pos_amt
            if month_key == current_month_calc:
                dealers_month[channel] += pos_amt
            # 首批订单月（取最早的月份）
            if channel not in dealers_first_month or month_key < dealers_first_month[channel]:
                dealers_first_month[channel] = month_key
        except:
            pass

    total_sales += pos_amt
    total_profit += profit
    total_orders.add(order_id)
    total_dealers.add(channel)

print(f"省区经理数: {len(managers)}")

# ========== 排序 ==========
month_keys = sorted(monthly.keys())
day_keys = sorted(daily.keys())
# 取最近14天的数据
recent_days = day_keys[-14:] if len(day_keys) > 14 else day_keys

dealer_ranking = sorted(dealers.items(), key=lambda x: x[1]['sales'], reverse=True)
product_ranking = sorted(products.items(), key=lambda x: x[1]['sales'], reverse=True)
manager_ranking = sorted(managers.items(), key=lambda x: x[1]['sales'], reverse=True)

# ========== 输出JSON数据给HTML用 ==========
# 本月数据
now = datetime.now()
current_month = now.strftime('%Y-%m')
month_sales = round(monthly[current_month]['sales'], 2) if current_month in monthly else 0
month_profit = round(monthly[current_month]['profit'], 2) if current_month in monthly else 0
month_orders = len(monthly[current_month]['orders']) if current_month in monthly else 0
month_profit_rate = round(month_profit/month_sales*100, 1) if month_sales > 0 else 0

# 今日数据（数据里最后一天）
today_sales = 0
today_orders = 0
today_key = ''
if len(day_keys) > 0:
    today_key = day_keys[-1]
    today_sales = daily[today_key]['sales']
    today_orders = len(daily[today_key]['orders'])

data = {
    'update_time': datetime.now().strftime('%Y-%m-%d %H:%M'),
    'total_sales': round(total_sales, 2),
    'total_profit': round(total_profit, 2),
    'total_profit_rate': round(total_profit/total_sales*100, 1) if total_sales > 0 else 0,
    'total_dealers': len(total_dealers),
    'total_orders': len(total_orders),
    'total_rows': len(filtered),
    # 本月
    'current_month': current_month,
    'month_sales': month_sales,
    'month_profit': month_profit,
    'month_orders': month_orders,
    'month_profit_rate': month_profit_rate,
    # 今日
    'today_key': today_key,
    'today_sales': round(today_sales, 2),
    'today_orders': today_orders,
    'months': month_keys,
    'monthly_data': [{
        'month': m,
        'sales': round(monthly[m]['sales'], 2),
        'profit': round(monthly[m]['profit'], 2),
        'orders': len(monthly[m]['orders'])
    } for m in month_keys],
    'daily_data': [{
        'day': d[5:], # 省略年份
        'sales': round(daily[d]['sales'], 2),
        'orders': len(daily[d]['orders'])
    } for d in recent_days],
    'managers': [{
        'name': m[0],
        'clients': len(m[1]['clients']),
        'today_sales': round(managers_today.get(m[0], 0), 2),
        'month_sales': round(managers_month[m[0]]['sales'], 2),
        'month_profit': round(managers_month[m[0]]['profit'], 2),
        'sales': round(m[1]['sales'], 2),
        'profit': round(m[1]['profit'], 2),
        'orders': len(m[1]['orders']),
        'profit_rate': round(m[1]['profit']/m[1]['sales']*100, 1) if m[1]['sales'] > 0 else 0
    } for m in manager_ranking],
    'dealers': [{
        'name': d[0],
        'manager': d[1]['manager'],
        'sales': round(d[1]['sales'], 2),
        'profit': round(d[1]['profit'], 2),
        'orders': len(d[1]['orders']),
        'qty': int(d[1]['qty']),
        'profit_rate': round(d[1]['profit']/d[1]['sales']*100, 1) if d[1]['sales'] > 0 else 0,
        'today_sales': round(dealers_today.get(d[0], 0), 2),
        'month_sales': round(dealers_month.get(d[0], 0), 2),
        'first_month': dealers_first_month.get(d[0], '-')
    } for d in dealer_ranking],
    'products': [{
        'name': p[0][:30],  # 截断长名
        'sales': round(p[1]['sales'], 2),
        'qty': int(p[1]['qty']),
        'profit': round(p[1]['profit'], 2),
        'avg_price': round(sum(p[1]['prices'])/len(p[1]['prices']), 2) if p[1]['prices'] else 0,
        'dealers': len(p[1]['dealers']),
        'profit_rate': round(p[1]['profit']/p[1]['sales']*100, 1) if p[1]['sales'] > 0 else 0
    } for p in product_ranking]
}

with open(r'E:\袋鼠先生\github-dashboard\dashboard_data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("\n数据已导出到 dashboard_data.json")
print(f"月度: {month_keys}")
