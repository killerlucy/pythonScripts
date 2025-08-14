import pandas as pd
import os

def merge_excel_files(base_dir, target_dir_name, output_file):
    # 初始化一个空的DataFrame用于存储合并后的数据
    combined_df = pd.DataFrame()
    
    # 遍历基础目录及其子目录
    for dirpath, dirnames, filenames in os.walk(base_dir):
        # 检查当前目录是否为名为target_dir_name的目录
        if os.path.basename(dirpath) == target_dir_name:
            # 遍历当前目录下的所有子目录
            for sub_dir in dirnames:
                sub_dir_path = os.path.join(dirpath, sub_dir)
                # 检查子目录下是否存在名为'aa.xlsx'的文件
                if '000001.xlsx' in os.listdir(sub_dir_path):
                    xlsx_path = os.path.join(sub_dir_path, '000001.xlsx')
                    print(f"正在处理文件：{xlsx_path}")
                    # 读取Excel文件
                    df = pd.read_excel(xlsx_path)
                    # 将读取的数据追加到combined_df中
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
    
    # 将合并后的数据保存到输出文件
    combined_df.to_excel(output_file, index=False)
    print(f"合并完成，输出文件已保存为：{output_file}")

# 调用函数，指定基础目录、目标目录名和输出文件名
base_directory = 'F:\\'  # F盘根目录
target_directory_name = '1403'  # 目标目录名
output_filename = 'merged_output.xlsx'  # 输出文件名
merge_excel_files(base_directory, target_directory_name, output_filename)