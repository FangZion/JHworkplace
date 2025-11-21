import pandas as pd
import random
from datetime import datetime, timedelta
import os

# 获取脚本所在目录
script_dir = os.path.dirname(os.path.abspath(__file__))

# 设置随机种子以保证结果可重现
random.seed(42)

# 公司列表
companies = [
    "众恒汽车部件有限公司",
    "嘉兴众恒汽车部件有限公司", 
    "安徽中恒电喷系统有限公司",
    "好事达国际（香港）有限公司",
    "众恒电喷系统（泰国）",
    "众恒国际（泰国）",
    "上海惠龙进出口贸易有限公司",
    "日昇瑞贸易",
    "嘉兴优科进出口有限公司",
    "浙江恒致精密制造有限公司",
    "上海天域联进出口贸易有限公司",
    "嘉兴佳展科技有限公司",
    "上海日昇瑞",
    "温州韦菲克科技有限公司"
]

# 回款类别列表
payment_types = ["电汇", "承兑汇票", "现金", "网银转账"]

# 币种列表
currencies = ["人民币", "USD", "EUR"]

# 汇率映射
exchange_rates = {"人民币": 1, "USD": 7.2, "EUR": 7.8}

# 生成200条随机数据
data = []
start_date = datetime(2025, 12, 1)
end_date = datetime(2025, 12, 31)

for i in range(200):
    # 随机日期（月内）
    random_date = start_date + timedelta(days=random.randint(0, 29))
    
    # 随机选择公司
    company = random.choice(companies)
    
    # 随机选择回款类别
    payment_type = random.choice(payment_types)
    
    # 随机选择币种
    currency = random.choice(currencies)
    
    # 随机原币金额（1000-50000之间，步长500）
    original_amount = random.randint(2, 100) * 500
    
    # 计算汇率和折本币金额
    exchange_rate = exchange_rates[currency]
    local_amount = original_amount * exchange_rate
    
    # 确定类型（根据公司名称判断）
    if any(x in company for x in ['众恒', '中恒']):
        record_type = "回款"
    else:
        record_type = "付款"
    
    data.append({
        "回款时间": random_date.strftime("%Y/%m/%d"),
        "回款类别": payment_type,
        "付款方名称": company,
        "币种": currency,
        "原币金额": original_amount,  # 保持数字格式
        "汇率": exchange_rate,
        "折本币": local_amount,  # 保持数字格式
        "类型": record_type
    })

# 创建DataFrame
df = pd.DataFrame(data)

# 按日期排序
df['回款时间'] = pd.to_datetime(df['回款时间'])
df = df.sort_values('回款时间')
df['回款时间'] = df['回款时间'].dt.strftime("%Y/%m/%d")

# 构建输出文件路径
output_path = os.path.join(script_dir, '12月交易明细.xlsx')

# 保存到Excel文件
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='交易明细', index=False)
    
    # 获取工作表并设置列宽和数字格式
    worksheet = writer.sheets['交易明细']
    column_widths = {
        'A': 12,  # 回款时间
        'B': 10,  # 回款类别
        'C': 25,  # 付款方名称
        'D': 8,   # 币种
        'E': 12,  # 原币金额
        'F': 8,   # 汇率
        'G': 12,  # 折本币
        'H': 8    # 类型
    }
    
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width
    
    # 设置数字格式（保留2位小数，千位分隔符）
    for row in range(2, len(df) + 2):  # 从第2行开始（跳过标题行）
        # 原币金额列（E列）
        worksheet[f'E{row}'].number_format = '#,##0.00'
        # 折本币列（G列）
        worksheet[f'G{row}'].number_format = '#,##0.00'
        # 汇率列（F列）
        worksheet[f'F{row}'].number_format = '0.0000'

print(f"Excel文件已生成到：{output_path}")
print(f"共生成 {len(df)} 条交易记录")
print("\n数据概览：")
print(f"- 时间范围：2025年12月1日 - 2025年12月31日")
print(f"- 涉及公司：{len(companies)} 家")
print(f"- 交易类型：回款/{len(df[df['类型']=='回款'])}条, 付款/{len(df[df['类型']=='付款'])}条")
print(f"- 币种分布：人民币/{len(df[df['币种']=='人民币'])}条, USD/{len(df[df['币种']=='USD'])}条, EUR/{len(df[df['币种']=='EUR'])}条")
print(f"- 金额格式：原币金额和折本币金额均为数字格式，可直接用于计算")