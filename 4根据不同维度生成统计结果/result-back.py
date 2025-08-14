# 程序说明：
# 批量处理一个文件夹下多个excel，每个excel表头相同，均为：
# 稽核规则名称，稽核资源ID，稽核资源名称，稽核失败时间，稽核资源结果，稽核失败天数，资源创建时间，省份，地市，厂家
# 1、添加一列名为”分类
# 2、筛选“稽核失败天数”字段，数值大于7。
# 3、"稽核规则名称"在自动采集稽核规则列表中的保留数据，不进行处理。
# 4、查询“资源创建时间”，时间格式均为：2024-09-14 08:27:59
# 5、判断“资源创建时间”在9月份之前的“分类”内容写：省份未及时处理。
# 6、判断“资源创建时间”在9月份之后的“分类”内容写：新建站未及时维护。
# 
import os
import pandas as pd
from datetime import datetime, timedelta

def get_judgement_date(audit_time):
    """根据稽核失败时间返回判断日期"""
    audit_datetime = pd.to_datetime(audit_time)
    year = audit_datetime.year
    month = audit_datetime.month
    
    # if month == 1:
    #     # 如果是1月，则向前退回到上一年的12月
    #     judgement_date = datetime(year - 1, 12, 23) - timedelta(days=1)
    # else:
    #     # 否则，将月份减1，并设置为12月23日
    #     judgement_date = datetime(year, month - 1, 23) - timedelta(days=1)

    # 返回当月1日0点的数据
    judgement_date = datetime(year, month , 1)
    # 返回标准化的日期（即去除时间部分）
    return judgement_date.replace(hour=0, minute=0, second=0)

def process_excel(file_path):
    df = pd.read_excel(file_path)
    
    # 转换日期时间格式
    df['稽核失败时间'] = pd.to_datetime(df['稽核失败时间'], errors='coerce')  # coerce将无效解析设为NaT
    df['资源创建时间'] = pd.to_datetime(df['资源创建时间'], errors='coerce')
    
    # 添加分类列
    df['分类'] = ''
    
    # 筛选告警次数大于7的数据
    filtered_df = df[df['稽核失败天数'] > 7]
    
    # 处理每一行数据
    for index, row in filtered_df.iterrows():
        judgement_date = get_judgement_date(row['稽核失败时间'])
        
        if pd.notna(row['资源创建时间']) and pd.notna(judgement_date):  # 防止NaN比较
            if row['资源创建时间'] < judgement_date:
                df.at[index, '分类'] = '存量资源未及时维护'
            else:
                df.at[index, '分类'] = '新建资源未及时维护'
    
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