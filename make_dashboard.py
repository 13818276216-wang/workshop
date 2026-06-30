# -*- coding: utf-8 -*-
import sys, json, math
sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from collections import defaultdict
from datetime import datetime

# ========== 读数据 ==========
path = r"E:\袋鼠先生\袋鼠先生MeOS\你的MeOS\01_我的资料\订单与交付\销售单明细账.xlsx"
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
    keywords = ['零食很忙', '赵一鸣', '鸣忙', '很忙零食']
    for kw in keywords:
        if kw in str(name):
            return True
    return False

filtered = [r for r in rows_data if not is_mingmang(r[col['客户名称']])]
print(f"剔除鸣鸣很忙后: {len(filtered)} 行")

# ========== 读合同台账获取省区经理 ==========
contract_path = r"E:\袋鼠先生\袋鼠先生MeOS\你的MeOS\01_我的资料\经销商管理\线下记录（授权书_证明关系）.xlsx"
try:
    wb_contract = openpyxl.load_workbook(contract_path, data_only=True)
    ws_contract = wb_contract['合同台账']
    contract_headers = [cell.value for cell in next(ws_contract.iter_rows(max_row=1))]
    contract_col = {h: i for i, h in enumerate(contract_headers)}
    
    manager_map = {} # 客户名称 -> 负责人
    for row in ws_contract.iter_rows(min_row=2, values_only=True):
        cname = str(row[contract_col.get('合作方', 2)] or '').strip()
        manager = str(row[contract_col.get('负责人', 11)] or '').strip()
        if cname and manager and manager != 'None':
            manager_map[cname] = manager
    wb_contract.close()
    print(f"成功读取合同台账，匹配到 {len(manager_map)} 个负责人的映射。")
except Exception as e:
    print("读取合同台账失败:", e)
    manager_map = {}

# ========== 解析数据 ==========
dealers = {}       # dealer_name -> {销售额, 毛利, 订单数, 负责人}
managers = {}      # manager_name -> {销售额, 毛利, 订单数, 负责的客户数}
products = {}      # product_name -> {销售额, 销量, 毛利, 单价列表}
monthly = defaultdict(lambda: {'sales': 0, 'profit': 0, 'orders': set()})
daily = defaultdict(lambda: {'sales': 0, 'orders': set()}) # 增加日维度
total_sales = 0
total_profit = 0
total_orders = set()
total_dealers = set()

cname_idx = col['客户名称']
pname_idx = col['货品名称']
qty_idx = col['数量']
price_idx = col['单价']
amt_idx = col['金额']
profit_idx = col['毛利']
time_idx = col['货品级发货时间']
order_idx = col['订单编号']

for r in filtered:
    cname = str(r[cname_idx] or '').strip()
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
    manager = manager_map.get(cname, '未知')
    
    if cname not in dealers:
        dealers[cname] = {'sales': 0, 'profit': 0, 'orders': set(), 'qty': 0, 'manager': manager}
    dealers[cname]['sales'] += pos_amt
    dealers[cname]['profit'] += profit
    dealers[cname]['orders'].add(order_id)
    dealers[cname]['qty'] += pos_qty

    if manager not in managers:
        managers[manager] = {'sales': 0, 'profit': 0, 'orders': set(), 'clients': set()}
    managers[manager]['sales'] += pos_amt
    managers[manager]['profit'] += profit
    managers[manager]['orders'].add(order_id)
    managers[manager]['clients'].add(cname)

    if pname not in products:
        products[pname] = {'sales': 0, 'qty': 0, 'profit': 0, 'prices': [], 'dealers': set()}
    products[pname]['sales'] += pos_amt
    products[pname]['qty'] += pos_qty
    products[pname]['profit'] += profit
    if price > 0:
        products[pname]['prices'].append(price)
    products[pname]['dealers'].add(cname)

    if time_str:
        try:
            month_key = time_str[:7]
            monthly[month_key]['sales'] += pos_amt
            monthly[month_key]['profit'] += profit
            monthly[month_key]['orders'].add(order_id)
            
            day_key = time_str[:10] # YYYY-MM-DD
            daily[day_key]['sales'] += pos_amt
            daily[day_key]['orders'].add(order_id)
        except:
            pass

    total_sales += pos_amt
    total_profit += profit
    total_orders.add(order_id)
    total_dealers.add(cname)

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
today_sales = 0
today_orders = 0
if len(day_keys) > 0:
    today_key = day_keys[-1] # 以数据里最后一天作为"今日"
    today_sales = daily[today_key]['sales']
    today_orders = len(daily[today_key]['orders'])

data = {
    'update_time': datetime.now().strftime('%Y-%m-%d %H:%M'),
    'total_sales': round(total_sales, 2),
    'total_profit': round(total_profit, 2),
    'total_dealers': len(total_dealers),
    'total_orders': len(total_orders),
    'total_rows': len(filtered),
    'profit_rate': round(total_profit/total_sales*100, 1) if total_sales > 0 else 0,
    'today_key': today_key if len(day_keys) > 0 else '',
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
        'sales': round(m[1]['sales'], 2),
        'profit': round(m[1]['profit'], 2),
        'orders': len(m[1]['orders']),
        'clients': len(m[1]['clients']),
        'profit_rate': round(m[1]['profit']/m[1]['sales']*100, 1) if m[1]['sales'] > 0 else 0
    } for m in manager_ranking],
    'dealers': [{
        'name': d[0],
        'manager': d[1]['manager'],
        'sales': round(d[1]['sales'], 2),
        'profit': round(d[1]['profit'], 2),
        'orders': len(d[1]['orders']),
        'qty': int(d[1]['qty']),
        'profit_rate': round(d[1]['profit']/d[1]['sales']*100, 1) if d[1]['sales'] > 0 else 0
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
