import pandas as pd
import os
from datetime import datetime

def split_and_save_files(base_directory, output_directory):
    # 获取目录下所有文件
    all_files = [f for f in os.listdir(base_directory) if f.endswith('.xlsx')]

    for file_name in all_files:
        file_path = os.path.join(base_directory, file_name)
        
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path, engine='openpyxl')

            # 提取源文件名的第一部分
            base_name_part = file_name.split('_')[0]
            
            # 根据稽核规则名称分组
            grouped = df.groupby('稽核规则名称')
            
            # 遍历每个分组并保存为新文件
            for rule_name, group_df in grouped:
                # 构建新文件名
                new_file_name = f"{base_name_part}_{rule_name}.xlsx"
                new_file_path = os.path.join(output_directory, new_file_name)

                # 保存分组数据到新文件
                group_df.to_excel(new_file_path, index=False, engine='xlsxwriter')
                print(f"Saved {new_file_path}")

        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

# 设置基础目录和输出目录
base_directory = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\稽核统计\统计数据'
output_directory = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\稽核统计\拆分结果'

# 确保输出目录存在
os.makedirs(output_directory, exist_ok=True)

print(f'Starting to process files in {base_directory}...')
split_and_save_files(base_directory, output_directory)
print('Processing completed.')