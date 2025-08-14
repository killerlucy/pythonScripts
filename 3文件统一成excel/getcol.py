import pandas as pd
import openpyxl
import os

# 文件路径
sample_file_path = 'D:/智网创新/无线网资源运营月报/月报ppt/10月/稽核详情测试/有源4G/4GRRU的收发模式完整性稽核_历史稽核数据详情 (1)/sftp/ads_zhw_wlzt_audit_result_detail_d_ss/北京市/audit_id=1405/day_id=20241001'

# 检查文件是否存在
if not os.path.exists(sample_file_path):
    print(f"File does not exist at path: {sample_file_path}")
else:
    try:
        # 使用 openpyxl 读取文件
        workbook = openpyxl.load_workbook(sample_file_path)
        sheet = workbook.active
        columns = [cell.value for cell in sheet[1]]
        print(columns)
    except PermissionError as pe:
        print(f"Permission denied error: {pe}")
    except FileNotFoundError as fnfe:
        print(f"File not found error: {fnfe}")
    except Exception as e:
        print(f"Error reading the file: {e}")

# 检查文件是否可读
if os.access(sample_file_path, os.R_OK):
    print(f"The file is readable.")
else:
    print(f"The file is not readable.")