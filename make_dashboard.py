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

# ========== 解析数据 ==========
dealers = {}       # dealer_name -> {销售额, 毛利, 订单数}
products = {}      # product_name -> {销售额, 销量, 毛利, 单价列表}
monthly = defaultdict(lambda: {'sales': 0, 'profit': 0, 'orders': set()})
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
    cname = str(r[cname_idx] or '')
    pname = str(r[pname_idx] or '')
    qty = float(r[qty_idx] or 0)
    price = float(r[price_idx] or 0)
    amt = float(r[amt_idx] or 0)
    profit = float(r[profit_idx] or 0) if r[profit_idx] is not None else 0
    order_id = str(r[order_idx] or '')
    time_str = str(r[time_idx] or '')

    if qty == 0 and amt == 0:
        continue

    # 只统计正数（排除退货/负数）
    pos_qty = max(qty, 0)
    pos_amt = max(amt, 0)
    # 毛利按实际（可正可负）
    
    if cname not in dealers:
        dealers[cname] = {'sales': 0, 'profit': 0, 'orders': set(), 'qty': 0}
    dealers[cname]['sales'] += pos_amt
    dealers[cname]['profit'] += profit
    dealers[cname]['orders'].add(order_id)
    dealers[cname]['qty'] += pos_qty

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
            month_key = time_str[:7]  # YYYY-MM
            monthly[month_key]['sales'] += pos_amt
            monthly[month_key]['profit'] += profit
            monthly[month_key]['orders'].add(order_id)
        except:
            pass

    total_sales += pos_amt
    total_profit += profit
    total_orders.add(order_id)
    total_dealers.add(cname)

print(f"经销商数: {len(dealers)}")
print(f"产品数: {len(products)}")
print(f"总销售额: {total_sales:,.0f}")
print(f"总毛利: {total_profit:,.0f}")
print(f"订单数: {len(total_orders)}")

# ========== 月度排序 ==========
month_keys = sorted(monthly.keys())

# ========== 经销商排名 ==========
dealer_ranking = sorted(dealers.items(), key=lambda x: x[1]['sales'], reverse=True)

# ========== 产品排名 ==========
product_ranking = sorted(products.items(), key=lambda x: x[1]['sales'], reverse=True)

# ========== 输出JSON数据给HTML用 ==========
data = {
    'update_time': datetime.now().strftime('%Y-%m-%d %H:%M'),
    'total_sales': round(total_sales, 2),
    'total_profit': round(total_profit, 2),
    'total_dealers': len(total_dealers),
    'total_orders': len(total_orders),
    'total_rows': len(filtered),
    'profit_rate': round(total_profit/total_sales*100, 1) if total_sales > 0 else 0,
    'months': month_keys,
    'monthly_data': [{
        'month': m,
        'sales': round(monthly[m]['sales'], 2),
        'profit': round(monthly[m]['profit'], 2),
        'orders': len(monthly[m]['orders'])
    } for m in month_keys],
    'dealers': [{
        'name': d[0],
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
