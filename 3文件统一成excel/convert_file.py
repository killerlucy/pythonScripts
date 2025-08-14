import pandas as pd
from pathlib import Path
import logging
import os
import tempfile

def setup_logging():
    # 使用系统临时目录保存日志文件
    log_file = 'file_processing.log'
    log_path = os.path.join(tempfile.gettempdir(), log_file)
    logging.basicConfig(filename=log_path, level=logging.INFO, 
                        format='%(asctime)s - %(levelname)s - %(message)s')
    print(f"Logging to {log_path}")

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
    
    encodings = ['utf-8', 'latin1', 'gbk', 'cp1252', 'utf-16']  # 尝试多种编码方式

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
                logging.info(f"Processed file {path} with encoding {encoding}.")
                processed = True
                break  # 成功处理后退出循环
            except UnicodeDecodeError as e:
                logging.error(f"Unicode decode error processing file {path} with encoding {encoding}: {e}")
                print(f"Unicode decode error processing file {path} with encoding {encoding}: {e}")
            except Exception as e:
                logging.error(f"Error processing file {path} with encoding {encoding}: {e}")
                print(f"Error processing file {path} with encoding {encoding}: {e}")
        
        if not processed:
            logging.error(f"Failed to process file {path} after trying all encodings.")
            print(f"Failed to process file {path} after trying all encodings.")

if __name__ == '__main__':
    setup_logging()
    directory_path = r'D:\智网创新\无线网资源运营月报\月报ppt\10月\省份稽核规则下载'  # 替换为你的Excel文件夹路径
    print('开始处理后缀名xlsx的unicodetxt文件,转换为真正的xlsx文件')
    convert_unicode_txt_to_xlsx(directory_path, '.xlsx')