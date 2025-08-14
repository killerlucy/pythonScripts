import pandas as pd
import os
from datetime import datetime
import warnings

# 定义映射表
rules_mapping = {
    '4GRRU关联机房放置点关联稽核': ('跨网络', '4G'),
     # 这个稽核规则改成行政区县了。暂时保留。
    '4GRRU经纬度所属行政区域准确性稽核': ('跨网络', '4G'),
    '4GRRU经纬度所属行政区县准确性稽核': ('跨网络', '4G'),
    '4GRRU经纬度完整性稽核': ('有源', '4G'),
    # 这个稽核规则名称后加了‘一致性稽核’。暂时保留。
    '4GRRU经纬度与所属安置地点经纬度': ('跨网络', '4G'),
    '4GRRU经纬度与所属安置地点经纬度一致性稽核': ('跨网络', '4G'),
    '4GRRU与天线关联稽核': ('无源', '4G'),
    '5GRRU经纬度完整性稽核': ('有源', '5G'),
    '5G无线网RRU与天线关联稽核': ('无源', '5G'),
     # 这个稽核规则改成行政区县了。暂时保留。
    'AAU经纬度所属行政区域准确性稽核': ('跨网络', '5G'),
    'AAU经纬度所属行政区县准确性稽核': ('跨网络', '5G'),
    'AAU经纬度与所属安置地点经纬度一致性稽核': ('跨网络', '5G'),
    'AAU收发模式完整性稽核': ('有源', '5G'),
    'AAU所属机房完整性稽核': ('跨网络', '5G'),
    'BBU关联机房放置点关联稽核': ('跨网络', '4G'),
    'CU关联所属机房完整性稽核': ('跨网络', '5G'),
    'DU关联机房完整性稽核': ('跨网络', '5G'),
    'EUTRANCELL关联RRU所属机房经纬度完整性稽核': ('跨网络', '4G'),
    #这个稽核名称后面加了：(不含中兴NB小区)
    'EUTRANCELL关联RRU所属机房经纬度完整性稽核(不含中兴NB小区)': ('跨网络', '4G'),
    'EUTRANCELL经纬度完整性稽核': ('有源', '4G'),
    'EUTRANCELL所属行政区域类型完整性稽核': ('有源', '4G'),
    'EUTRANCELL小区覆盖类型完整性稽核': ('有源', '4G'),
    'GNODEB所属行政区域完整性稽核': ('有源', '5G'),
    'NRCELLDU_关联AAU_关联稽核': ('有源', '5G'),
    'NRCELLDU工作频段完整性稽核': ('有源', '5G'),
    'NRCELLDU关联AAU所属机房经纬度完整性稽核': ('跨网络', '5G'),
    'NRCELLDU所属行政区域类型完整性稽核': ('有源', '5G'),
    'NRCELLDU下行频点完整性稽核': ('有源', '5G'),
    'NRCELLDU小区覆盖类型完整性稽核': ('有源', '5G'),
    '当日-无线专业-4G-ENODEB-资源与告警关联率': ('跨域', '4G'),
    '当日-无线专业-5G-GNODEB-资源与告警关联率': ('跨域', '5G'),
    '联通5G天线电子下倾角完整性稽核': ('无源', '5G'),
    '联通5G天线机械倾角完整性稽核': ('无源', '5G'),
    '联通NRCELLDU所属行政区域完整性稽核': ('有源', '5G'),
    '铁塔站址编码匹配率': ('跨域', '设备'),
    '无线网室外物理站址距离合规性稽核': ('跨网络', '设备'),
    # 这个稽核规则改成行政区县了。暂时保留。
    '4G小区经纬度所属行政区域准确性稽核': ('跨网络', '4G'),
    '4G小区经纬度所属行政区县准确性稽核': ('跨网络', '4G'),
    'BBUCUDU机房放置点经纬度所属行政区域准确性稽核': ('跨网络', '4/5G'),
    'RRUAAU机房放置点经纬度所属行政区域准确性稽核': ('跨网络', '4/5G'),
    '机房铁塔站址编码完整性稽核': ('跨域', '设备'),
    '设备室外放置点铁塔站址编码完整性稽核': ('跨域', '设备'),
    '联通4G天线挂高完整性稽核': ('无源', '4G'),
    '5G无线网AAU与天线关联稽核': ('无源', '5G'),
    '联通5G天线方向角完整性稽核': ('无源', '5G'),
    '联通5G天线挂高完整性稽核': ('无源', '5G'),
    '4GRRU的收发模式完整性稽核': ('有源', '4G'),
    '4G基站所属管理区域完整性稽核': ('有源', '4G'),
    'ENODEB所属行政区域完整性稽核': ('有源', '4G'),
    'ENODEB与BBU关联稽核': ('有源', '4G'),
    'ENODEB子网掩码完整性稽核': ('有源', '4G'),
    'EUTRANCELL下行频点完整性稽核': ('有源', '4G'),
    '联通EUTRANCELL工作频段完整性稽核': ('有源', '4G'),
    '联通EUTRANCELL所属行政区域完整性稽核': ('有源', '4G'),
    'CU是否关联到所属的GNODEB基站': ('有源', '5G'),
    'GNODEB管理区县完整性稽核': ('有源', '5G')
}

def get_province_from_path(path):
    # 获取路径中倒数第三个目录的名称
    path_parts = path.split(os.sep)
    if len(path_parts) >= 3:
        return path_parts[-3]
    else:
        return None

def process_and_merge_files(base_directory, filter_file):
    combined_df = pd.DataFrame()
    required_columns = ['省份', '分类', '网络类型']

    # 忽略 UserWarning
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

    for dirpath, _, filenames in os.walk(base_directory):
        for filename in filenames:
            if '~$' in filename or filename.startswith('.'):
                continue
            
            file_path = os.path.join(dirpath, filename)
            
            if filename == '000001.xlsx':
                province_name = get_province_from_path(dirpath)

                try:
                    if not os.access(file_path, os.R_OK):
                        print(f"File {file_path} is not readable.")
                        continue
                    
                    columns_to_read = [
                        '稽核规则id', '稽核资源ID', '稽核规则名称', '稽核资源名称',
                        '稽核资源归属地市', '稽核资源归属区县', '创建时间',
                        '资源创建时间', '失败原因'
                    ]
                    
                    df = pd.read_excel(file_path, usecols=lambda x: x in columns_to_read, engine='openpyxl')
                    
                    df['省份'] = province_name
                    df['分类'] = df['稽核规则名称'].map(lambda x: rules_mapping.get(x, ('', ''))[0])
                    df['网络类型'] = df['稽核规则名称'].map(lambda x: rules_mapping.get(x, ('', ''))[1])

                    rename_dict = {
                        '稽核资源归属地市': '地市',
                        '稽核资源归属区县': '区县'
                    }
                    df.rename(columns=rename_dict, inplace=True)

                    missing_columns = [col for col in required_columns if col not in df.columns]
                    if missing_columns:
                        print(f"Missing columns in {file_path}: {missing_columns}")
                        continue

                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                
                except PermissionError as pe:
                    print(f"Permission denied error on {file_path}: {pe}")
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")
                    continue

    missing_columns = [col for col in required_columns if col not in combined_df.columns]
    if missing_columns:
        print(f"Missing columns in combined DataFrame: {missing_columns}")
        return

    fixed_columns = ['省份', '分类', '网络类型', '稽核规则id', '稽核规则名称']
    columns_order = fixed_columns + [col for col in combined_df.columns if col not in fixed_columns]
    combined_df = combined_df[columns_order]

    # 去除重复的记录，保留每种组合每天第一次出现的记录
    failure_days = combined_df.drop_duplicates(subset=['省份', '稽核规则名称', '稽核资源ID'], keep='first')

    # 保存最终结果到Excel文件中
    failure_days.to_excel(filter_file, index=False, engine='xlsxwriter')


# 获取当前时间并格式化
now = datetime.now()
time_str = now.strftime("%Y%m%d_%H%M%S")

# 设置基础目录和输出文件名
base_directory = r'D:\智网创新\2月报\月报ppt\12月\12月派单分析数据\派单量大稽核统计\31省稽核数据\4G小区经纬度所属行政区县准确性稽核\sftp\ads_zhw_wlzt_audit_result_detail_d_ss'
filter_filename = os.path.join(r'D:\智网创新\2月报\月报ppt\12月\12月派单分析数据\派单量大稽核统计\派单量大稽核统计结果', f'4G小区经纬度所属行政区县准确性稽核{time_str}.xlsx')

print(f'Processing {base_directory}...')
print(f'Processing Result file : {filter_filename}...')
process_and_merge_files(base_directory, filter_filename)
process_and_merge_files(base_directory, filter_filename)