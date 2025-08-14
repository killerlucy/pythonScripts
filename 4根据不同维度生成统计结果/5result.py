# 程序说明：
# 批量处理一个文件夹下多个excel，每个excel表头相同，均为：
# 省份	地市	分类	网络类型	稽核规则id	稽核规则名称	稽核资源ID	稽核资源名称	资源创建时间	稽核失败时间	稽核失败日期	失败原因	稽核失败天数
# 1、添加一列名为：分析
# 3、"稽核规则名称"在自动采集稽核规则列表中的保留数据不进行处理，不在的继续完成以下程序
# 4、查询“资源创建时间”，时间格式均为：2024-09-14 08:27:59
# 5、判断“资源创建时间”在11月份之前的“分类”内容写：存量资源。
# 6、判断“资源创建时间”在11月份之后的“分类”内容写：新增资源。

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

def process_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        
        # 转换日期时间格式
        df['稽核失败时间'] = pd.to_datetime(df['稽核失败时间'], errors='coerce')
        df['资源创建时间'] = pd.to_datetime(df['资源创建时间'], errors='coerce')
        
        # 添加维护方式列
        df['维护方式'] = df['稽核规则名称'].apply(lambda x: '自动采集' if x in excluded_rules else '手动维护')

        # 添加分析列
        df['分析'] = ''  # 使用“分析”作为列名
        
        # 定义存量资源和新增资源的分界线-----------------------此处需要改日期--------------------
        cutoff_date = pd.Timestamp(2025, 1, 1)

        # 处理每一行数据
        for index, row in df.iterrows():
            if not pd.isna(row['资源创建时间']):
                if row['资源创建时间'] < cutoff_date:
                    df.at[index, '分析'] = '存量资源'
                else:
                    df.at[index, '分析'] = '新增资源'

        # 去重：基于特定列组合，保留第一次出现的数据
        df.drop_duplicates(subset=['省份','稽核规则名称','稽核资源ID'], keep='first', inplace=True)

        # 重新排序列，确保“分析”列在最后
        columns_order = ['省份', '地市', '区县','分类', '网络类型', '稽核规则id', '稽核规则名称', '稽核资源ID', '稽核资源名称', 
                         '资源创建时间', '稽核失败时间', '稽核失败日期', '失败原因', '稽核失败天数', '维护方式', '分析']
        df = df.reindex(columns=columns_order)
        
        # 输出处理后的数据到新的Excel文件
        output_file_path = file_path.replace('.xlsx', '_processed.xlsx')
        # 避免覆盖已经处理过的文件
        if not os.path.exists(output_file_path):
            df.to_excel(output_file_path, index=False)
        else:
            print(f"文件 {output_file_path} 已存在，跳过处理.")

    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")

def batch_process(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx') and not filename.endswith('_processed.xlsx'):
            file_path = os.path.join(directory, filename)
            process_excel(file_path)

if __name__ == "__main__":
    # ---------------------替换为你的Excel文件夹路径-------------------------
    directory = r'D:\智网创新\2月报\月报ppt\1月\1月月报失败明细\稽核统计'  
    # directory = r'D:\智网创新\2月报\月报ppt\12月\12月派单分析数据\派单量大稽核统计\派单量大稽核统计结果\one'  
    batch_process(directory)