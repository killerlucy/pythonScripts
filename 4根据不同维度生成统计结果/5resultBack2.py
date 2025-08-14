import os
import pandas as pd
from datetime import datetime

# 自动采集稽核规则列表
excluded_rules = [
    '4G小区与RRU关联稽核',
    '联通EUTRANCELL工作频段完整性稽核',
    'EUTRANCELL下行频点完整性稽核',
    'EUTRANCELL关联所属基站完整性稽核',
    '4GRRU的收发模式完整性稽核',
    'ENODEB与BBU关联稽核',
    'ENODEBIP地址完整性稽核',
    'ENODEB子网掩码完整性稽核',
    'NRCELLDU关联所属基站完整性稽核',
    'NRCELLDU_关联AAU_关联稽核',
    'NRCELLDU工作频段完整性稽核',
    'NRCELLDU下行频点完整性稽核',
    'AAU收发模式完整性稽核',
    'AAU经纬度完整性稽核',
    'GNODEB子网掩码完整性稽核',
    'GNODEBIP地址完整性稽核',
    'CU是否关联到所属的GNODEB基站',
    '5G无线网AAU与天线关联稽核',
    '当日-无线专业-4G-ENODEB-资源与告警关联率',
    '当日-无线专业-5G-GNODEB-资源与告警关联率'
]

def get_judgement_date(audit_time):
    """根据稽核失败时间返回判断日期"""
    # 尝试转换为日期格式
    audit_datetime = pd.to_datetime(audit_time, errors='coerce')
    
    year = audit_datetime.year
    month = audit_datetime.month
    
    # 返回当月1日0点的数据
    judgement_date = datetime(year, month, 1)
    # 返回标准化的日期（即去除时间部分）
    return judgement_date.replace(hour=0, minute=0, second=0)

def process_excel(file_path):
    df = pd.read_excel(file_path)
    
    # 转换日期时间格式
    df['稽核失败时间'] = pd.to_datetime(df['稽核失败时间'], errors='coerce')  # coerce将无效解析设为NaT
    
    # 支持多种日期时间格式
    def parse_date(date_str):
        formats = ['%Y%m%d', '%Y-%m-%d %H:%M:%S']
        for fmt in formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return pd.NaT
    
    df['资源创建时间'] = df['资源创建时间'].apply(lambda x: parse_date(x) if isinstance(x, str) else pd.NaT)
    
    # 添加分类列
    df['分类'] = ''
    
    # 排除特定规则名称
    filtered_df = df[~df['稽核规则名称'].isin(excluded_rules)]
    
    # 处理每一行数据
    for index, row in filtered_df.iterrows():
        if pd.isna(row['资源创建时间']):
            continue  # 如果资源创建时间是NaN，则跳过该行
        
        # 获取当月第一天
        judgement_date = get_judgement_date(row['资源创建时间'])
        
        if judgement_date.year < 2024 or (judgement_date.year == 2024 and judgement_date.month < 9):
            df.at[index, '分类'] = '省份未及时处理'
        else:
            df.at[index, '分类'] = '新建站未及时维护'
    
    # 重新排序列
    columns_order = ['省份', '地市', '稽核规则名称', '稽核资源ID', '稽核资源名称', '稽核失败时间', '稽核资源结果', '稽核失败天数', '资源创建时间', '厂家', '分类']
    df = df.reindex(columns=columns_order)
    
    # 输出处理后的数据到新的Excel文件
    output_file_path = file_path.replace('.xlsx', '_processed.xlsx')
    df.to_excel(output_file_path, index=False)

def batch_process(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(directory, filename)
            process_excel(file_path)

if __name__ == "__main__":
    directory = 'D:/智网创新/无线网资源运营月报/月报ppt/9月/9月稽核明细统计结果4类/3result'  # 替换为你的Excel文件夹路径
    batch_process(directory)