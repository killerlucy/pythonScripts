import os
import pandas as pd

# 主Excel文件路径
main_file_path = r'D:\智网创新\故障派单失败\Fw_故障中心平台运营质量中资源数据问题，请配合支撑，谢谢！\附件二各专业派单失败详单-20250124\各专业派单失败详单-20250124\无线网.xlsx'
# 存放其他Excel文件的文件夹路径
folder_path = r'D:\智网创新\故障派单失败\Fw_故障中心平台运营质量中资源数据问题，请配合支撑，谢谢！\附件二各专业派单失败详单-20250124\各专业派单失败详单-20250124\匹配信息'

# 读取主Excel文件
try:
    main_df = pd.read_excel(main_file_path)
except Exception as e:
    print(f"读取主Excel文件时发生错误: {e}")
    exit(1)

# 确保主文件中存在所有必要的列
if 'create_date' not in main_df.columns:
    main_df['create_date'] = None
if '判断结果' not in main_df.columns:
    main_df['判断结果'] = None
if 'delete_time' not in main_df.columns:
    main_df['delete_time'] = None
if 'p_confirm_status' not in main_df.columns:
    main_df['p_confirm_status'] = None

# 打印主数据帧的列名，以确认是否已添加
print("更新后的主数据帧列名:", main_df.columns)

# 获取文件夹内所有Excel文件
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# 遍历文件夹内的每个Excel文件
for file in files:
    try:
        # 读取当前Excel文件
        df_temp = pd.read_excel(os.path.join(folder_path, file))
    except Exception as e:
        print(f"读取文件 {file} 时发生错误: {e}")
        continue
    
    # 打印临时数据帧的列名，以确认是否包含所有必需的列
    print(f"文件 {file} 的列名:", df_temp.columns)
    
    # 确保临时数据帧中有我们需要的列
    required_columns = ['create_date', 'p_confirm_status', 'delete_time', '判断结果']
    if not all(col in df_temp.columns for col in required_columns):
        print(f"警告: 文件 {file} 缺少必要的列。")
        continue
    
    # 遍历主数据帧中的每一行，根据设备类型选择匹配列
    for index, row in main_df.iterrows():
        device_type = row['设备类型']
        match_col = None
        if device_type == 'ENodeB':
            if 'EMS_ORIG_RES_ID' in df_temp.columns:
                match_col = 'EMS_ORIG_RES_ID'
            else:
                print(f"警告: 文件 {file} 不包含 EMS_ORIG_RES_ID 列。")
        elif device_type == 'GNodeB':
            if 'NMS_ORIG_RES_ID' in df_temp.columns:
                match_col = 'NMS_ORIG_RES_ID'
            else:
                print(f"警告: 文件 {file} 不包含 NMS_ORIG_RES_ID 列。")
        else:
            print(f"警告: 未知设备类型 {device_type}。")
        
        if match_col is None:
            print(f"警告: 文件 {file} 不包含适合当前设备类型的匹配列。")
            continue
        
        # 去除空格和特殊字符
        resource_id = row['资源ID'].strip()
        df_temp[match_col] = df_temp[match_col].str.strip()

        # 根据资源ID进行数据匹配
        matched_row = df_temp[df_temp[match_col] == resource_id]
        if not matched_row.empty:
            main_df.at[index, 'create_date'] = matched_row['create_date'].values[0]
            main_df.at[index, '判断结果'] = matched_row['判断结果'].values[0]
            main_df.at[index, 'delete_time'] = matched_row['delete_time'].values[0]
            main_df.at[index, 'p_confirm_status'] = matched_row['p_confirm_status'].values[0]
            print(f"资源ID {resource_id} 在文件 {file} 中匹配成功。")
        else:
            print(f"资源ID {resource_id} 在文件 {file} 中未匹配成功。")

# 在最后再次打印主数据帧的列名，以确认是否已正确更新
print("最终的主数据帧列名:", main_df.columns)

# 将时间列转换为字符串格式
time_columns = ['create_date', 'delete_time']
for col in time_columns:
    if col in main_df.columns:
        main_df[col] = pd.to_datetime(main_df[col], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')

# 保存更新后的主Excel文件
try:
    main_df.to_excel(main_file_path, index=False)
    print("更新后的主Excel文件已保存。")
except Exception as e:
    print(f"保存更新后的主Excel文件时发生错误: {e}")