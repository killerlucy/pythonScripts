import pandas as pd
from pathlib import Path
import os
import shutil

def convert_unicode_txt_to_xlsx(directory, suffix):
    root_path = Path(directory)
    file_suffix = suffix
    file_mapping = '*' + file_suffix
    fixed_columns = [
        '主键', '稽核规则id', '稽核规则名称', '稽核资源ID', '稽核资源名称', 
        '稽核资源归属区县ID', '稽核资源归属区域ID', '资源类型ID', '稽核资源结果', 
        '规则执行批次', '用户可查看稽核结果日期', '稽核结果描述', '删除状态', 
        '创建时间', '结果表日期分区字段', '失败原因', '失败原因ID'
    ]
    
    encodings = ['utf-8', 'latin1', 'gbk']  # 尝试多种编码方式
    log_file = 'failed_files.log'

    with open(log_file, 'w') as log:
        for path in root_path.rglob(file_mapping):
            file_name = path.name
            print('processing file:', path)
            base_name = file_name.split('.')[0]
            output_file = path.with_name(base_name + '.xlsx')
            processed = False
            
            for encoding in encodings:
                try:
                    # 读取文件头部，获取列数
                    with open(path, 'r', encoding=encoding) as f:
                        first_line = f.readline().strip()
                        num_columns = len(first_line.split('\t'))
                    
                    # 动态生成列名
                    if num_columns < len(fixed_columns):
                        common_columns = fixed_columns[:num_columns]
                    else:
                        common_columns = fixed_columns + [f'拓展字段{i}' for i in range(1, num_columns - len(fixed_columns) + 1)]
                    
                    df = pd.read_csv(path, sep='\t', encoding=encoding, engine='python', header=None, skiprows=[0])
                    df.columns = common_columns
                    df.to_excel(output_file, index=False)
                    print(output_file, 'processed finished.')
                    processed = True
                    break  # 成功处理后退出循环
                except Exception as e:
                    print(f"Error processing file {path} with encoding {encoding}: {e}")
                    continue
            
            if not processed:
                log.write(f"Failed to process file {path}\n")

def replace_string_in_csv(file_path):
    temp_file_path = file_path + '.tmp'
    lines = []
    with open(file_path, 'r', newline='', encoding='utf-8') as infile:
        for line in infile:
            modified_line = line.replace('|+|', '\t')
            if modified_line.startswith('"') and modified_line.endswith('"'):
                modified_line = modified_line[1:-1]
            lines.append(modified_line)
    with open(temp_file_path, 'w', newline='', encoding='utf-8') as outfile:
        for line in lines:
            outfile.write(line)
    os.remove(file_path)
    os.rename(temp_file_path, file_path)

def process_directory(directory_path):
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.csv'):
                file_path = os.path.join(root, file)
                print(f'Processing {file_path}...')
                replace_string_in_csv(file_path)

def remove_directories(root_dir, dir_names):
    for dir_name in dir_names:
        for root, dirs, files in os.walk(root_dir):
            if dir_name in dirs:
                path_to_remove = os.path.join(root, dir_name)
                try:
                    shutil.rmtree(path_to_remove)
                    print(f"Directory {path_to_remove} has been deleted.")
                except Exception as e:
                    print(f"Failed to delete directory {path_to_remove}: {e}")

if __name__ == '__main__':
    directory_path = r'D:\智网创新\无线网资源运营月报\月报ppt\10月\省份稽核规则下载'  # 替换为你的Excel文件夹路径
    # directories_to_delete=['day_id=20240901','day_id=20240902','day_id=20240903','day_id=20240904','day_id=20240905','day_id=20240906','day_id=20240907','day_id=20240908','day_id=20240909','day_id=20240910','day_id=20240911','day_id=20240912']
    # remove_directories(directory_path, directories_to_delete)
    # print('开始处理后缀名csv文件“|+|”转换为制表符')
    # process_directory(directory_path)
    print('开始处理后缀名xlsx的unicodetxt文件,转换为真正的xlsx文件')
    convert_unicode_txt_to_xlsx(directory_path, '.xlsx')
    # print('开始处理后缀名csv文件,转换为真正的xlsx文件')
    # convert_unicode_txt_to_xlsx(directory_path, '.csv')