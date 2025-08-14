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
                        '资源创建时间', '失败原因',
                        '设备创建时间','承建方GNODEB ID','资源场创建时间','小区创建时间'
                    ]
                    
                    df = pd.read_excel(file_path, usecols=lambda x: x in columns_to_read, engine='openpyxl')
                    
                    df['省份'] = province_name
                    df['分类'] = df['稽核规则名称'].map(lambda x: rules_mapping.get(x, ('', ''))[0])
                    df['网络类型'] = df['稽核规则名称'].map(lambda x: rules_mapping.get(x, ('', ''))[1])

                    creation_time_cols = [col for col in df.columns if '创建时间' in col]
                    if len(creation_time_cols) == 1:
                        df.rename(columns={creation_time_cols[0]: '稽核失败时间'}, inplace=True)
                        df['稽核失败时间'] = pd.to_datetime(df['稽核失败时间'], errors='coerce')
                        df['稽核失败时间日期'] = df['稽核失败时间'].dt.strftime('%Y-%m-%d')
                    elif len(creation_time_cols) > 1:
                        df.rename(columns={creation_time_cols[0]: '稽核失败时间', creation_time_cols[1]: '资源创建时间'}, inplace=True)
                        df['稽核失败时间'] = pd.to_datetime(df['稽核失败时间'], errors='coerce')
                        df['稽核失败时间日期'] = df['稽核失败时间'].dt.strftime('%Y-%m-%d')
                        df['资源创建时间'] = pd.to_datetime(df['资源创建时间'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
                        df['资源创建时间日期'] = df['资源创建时间'].dt.strftime('%Y-%m-%d')

                    # 删除多余 '创建时间' 列
                    for col in creation_time_cols[2:]:
                        if col in df.columns:
                            df.drop(columns=[col], inplace=True)

                    # 统一其他类似列名为 "资源创建时间"
                    for col in ['设备创建时间', '承建方GNODEB ID', '资源场创建时间', '小区创建时间']:
                        if col in df.columns and '资源创建时间' not in df.columns:
                            df.rename(columns={col: '资源创建时间'}, inplace=True)
                        elif col in df.columns and '资源创建时间' in df.columns:
                            df.drop(columns=[col], inplace=True)

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

    # 去除每天重复的记录，保留每种组合每天第一次出现的记录
    daily_unique_failures = combined_df.drop_duplicates(subset=['稽核规则名称', '稽核资源ID', '稽核失败时间日期'], keep='first')

    # 计算每个规则和资源ID组合的不同日期数量（即失败天数）
    failure_days = daily_unique_failures.groupby(['稽核规则名称', '稽核资源ID'])['稽核失败时间日期'].nunique().reset_index(name='稽核失败天数')

    # 将失败天数合并回原始数据框
    combined_df = combined_df.merge(failure_days, on=['稽核规则名称', '稽核资源ID'], how='left')

    # 如果需要，可以再次去重以保留每种组合最早的记录
    first_occurrences = combined_df.drop_duplicates(subset=['稽核规则名称', '稽核资源ID'], keep='first')

    # 保存最终结果到Excel文件中
    first_occurrences.to_excel(filter_file, index=False, engine='xlsxwriter')

# 获取当前时间并格式化
now = datetime.now()
time_str = now.strftime("%Y%m%d_%H%M%S")


# 设置基础目录和输出文件名
# rule_name = '跨网络4G'
# fresult_name = '4GRRU关联机房放置点关联稽核'
# fresult_name = '4GRRU经纬度所属行政区域准确性稽核'
# fresult_name = '4GRRU经纬度与所属安置地点经纬度'
# fresult_name = 'BBU关联机房放置点关联稽核'
# fresult_name = 'EUTRANCELL关联RRU所属机房经纬度完整性稽核'
# rule_name = '跨网络5G'
# fresult_name = 'AAU经纬度与所属安置地点经纬度一致性稽核'
# fresult_name = 'AAU经纬度所属行政区域准确性稽核'
# fresult_name = 'AAU所属机房完整性稽核'
# fresult_name = 'CU关联所属机房完整性稽核'
# fresult_name = 'DU关联机房完整性稽核'
# fresult_name = 'NRCELLDU关联AAU所属机房经纬度完整性稽核'
# rule_name = '跨网络设备'
# fresult_name = '无线网室外物理站址距离合规性稽核'
# rule_name = '跨域'
# fresult_name = '当日-无线专业-4G-ENODEB-资源与告警关联率'
# fresult_name = '当日-无线专业-5G-GNODEB-资源与告警关联率'
# fresult_name = '铁塔站址编码匹配率'
# rule_name = '无源4G'
# fresult_name = '4GRRU与天线关联稽核'
# rule_name = '无源5G'
# fresult_name = '5G无线网RRU与天线关联稽核'
# fresult_name = '联通5G天线电子下倾角完整性稽核'
# fresult_name = '联通5G天线机械倾角完整性稽核'
# rule_name = '有源4G'
# fresult_name = '4GRRU经纬度完整性稽核'
# fresult_name = 'EUTRANCELL经纬度完整性稽核'
# fresult_name = 'EUTRANCELL所属行政区域类型完整性稽核'
# fresult_name = 'EUTRANCELL小区覆盖类型完整性稽核'
# rule_name = '有源5G'
# fresult_name = 'AAU收发模式完整性稽核'
# fresult_name = 'GNODEB所属行政区域完整性稽核'
# fresult_name = 'NRCELLDU_关联AAU_关联稽核'
# fresult_name = 'NRCELLDU工作频段完整性稽核'
# fresult_name = 'NRCELLDU所属行政区域类型完整性稽核'
# fresult_name = 'NRCELLDU下行频点完整性稽核'
# fresult_name = 'NRCELLDU小区覆盖类型完整性稽核'
# fresult_name = '联通NRCELLDU所属行政区域完整性稽核'
# 统计报表的汇总如下：匹配资源创建时间、省份、地市，使用合并文件
rule_name = '跨网络合并'
fresult_name = '跨网络45G经纬度'
# fresult_name = '跨网络4G'
# fresult_name = '跨网络5G'
# fresult_name = '跨网络设备'
# rule_name = '跨域'
# fresult_name = '跨域'
# rule_name = '有源合并'
# fresult_name = '有源45G'
# rule_name = '无源合并'
# fresult_name = '无源45G'
source_file_path = os.path.join(rule_name, fresult_name)

base_directory = r'D:\智网创新\2月报\月报ppt\1月\1月月报失败明细\\' + source_file_path
filter_filename = os.path.join(r'D:\智网创新\2月报\月报ppt\1月\1月月报失败明细\稽核统计', fresult_name + '_统计' + time_str + '.xlsx')

print(f'Processing {base_directory}...')
print(f'Processing Result file : {filter_filename}...')
process_and_merge_files(base_directory, filter_filename)