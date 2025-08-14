import pandas as pd
import os
from datetime import datetime

# 定义文件夹路径
folder_path = r'D:\智网创新\故障派单失败\11月\分月稽核派单统计2024年11月25日'
output_folder = r'D:\智网创新\故障派单失败\11月\分月稽核派单统计2024年11月25日\合并结果'

# 创建输出文件夹（如果不存在）
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# 获取文件夹下所有的Excel文件
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# 初始化一个空的DataFrame用于存储最终的结果
final_df = pd.DataFrame(columns=['省份', '稽核规则名称'])

# 遍历每一个文件
for file in files:
    # 从文件名中提取月份
    month = int(file.split('月')[0])
    
    # 读取Excel文件
    df = pd.read_excel(os.path.join(folder_path, file))
    
    # 去重
    df = df.drop_duplicates(subset=['区域', '规则名称'])
    
    # 固定列信息
    fixed_columns = df[['区域', '规则名称']].copy()
    fixed_columns.columns = ['省份', '稽核规则名称']
    
    # 动态列信息
    dynamic_columns = df[['待处理', '核对成功', '七日内核对成功数量']].copy()
    dynamic_columns[f'{month}月异常数据处理量'] = dynamic_columns['核对成功']
    dynamic_columns[f'{month}月派单量'] = dynamic_columns['待处理'] + dynamic_columns['核对成功']
    dynamic_columns[f'{month}月闭环率'] = dynamic_columns['核对成功'] / (dynamic_columns['核对成功'] + dynamic_columns['待处理'])
    dynamic_columns[f'{month}月异常数据7日处理量'] = dynamic_columns['七日内核对成功数量']
    dynamic_columns[f'{month}月处理及时率'] = dynamic_columns['七日内核对成功数量'] / dynamic_columns['核对成功']
    
    # 保留需要的列
    temp_df = pd.concat([fixed_columns, dynamic_columns[[f'{month}月异常数据处理量', f'{month}月派单量', f'{month}月闭环率', f'{month}月异常数据7日处理量', f'{month}月处理及时率']]], axis=1)
    
    # 合并固定列和动态列
    if final_df.empty:
        final_df = temp_df
    else:
        # 使用唯一的后缀避免列名冲突
        final_df = pd.merge(final_df, temp_df, on=['省份', '稽核规则名称'], how='outer', suffixes=(f'_x', f'_y'))

# 去除后缀
final_df.columns = [col.split('_')[0] for col in final_df.columns]

# 对包含“率”的列进行格式化
rate_columns = [col for col in final_df.columns if '率' in col]
for col in rate_columns:
    final_df[col] = final_df[col].apply(lambda x: '{:.2%}'.format(x) if pd.notnull(x) else '')

# 按照指定的顺序输出
columns_order = ['省份', '稽核规则名称']

# 生成1月到12月的列顺序
for month in range(1, 13):
    columns_order.extend([
        f'{month}月异常数据处理量',
        f'{month}月派单量',
        f'{month}月闭环率',
        f'{month}月异常数据7日处理量',
        f'{month}月处理及时率'
    ])

# 过滤掉不存在的列
final_df = final_df[final_df.columns.intersection(columns_order)]

# 保存最终结果
current_time = datetime.now().strftime('%Y-%m-%d%H%M%S')
output_file_name = f'异常数据合并{current_time}.xlsx'
output_path = os.path.join(output_folder, output_file_name)
final_df.to_excel(output_path, index=False)

print(f"文件已保存至 {output_path}")