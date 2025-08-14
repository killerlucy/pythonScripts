import os
import pandas as pd
import logging

# 设置日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义目录路径
dir_a = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\统计数据\无源'
dir_b = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\分裂后的数据'

# 创建一个空列表来存储目录B中所有文件的数据
data_b = []

# 读取目录B中的所有Excel文件，并存储相关列的数据
logging.info("开始读取目录B中的Excel文件...")
for file in os.listdir(dir_b):
    if file.endswith('.xlsx') or file.endswith('.xls'):
        filepath = os.path.join(dir_b, file)
        try:
            df = pd.read_excel(filepath)
            # 检查列名是否存在
            if 'eqp_id' not in df.columns:
                logging.error(f"文件 {filepath} 缺少 'eqp_id' 列")
                continue
            
            # 将eqp_id转换为文本格式
            df['eqp_id'] = df['eqp_id'].astype(str)
            required_columns = ['eqp_id', 'create_date', 'region_name']
            if not all(col in df.columns for col in required_columns):
                missing_cols = [col for col in required_columns if col not in df.columns]
                logging.error(f"文件 {filepath} 缺少必要列：{missing_cols}")
                continue
            
            data_b.append(df[required_columns])
        except Exception as e:
            logging.error(f"无法读取文件 {filepath}: {e}")

# 如果目录B中有多个文件，将它们的数据合并成一个大表
if len(data_b) > 1:
    b_df = pd.concat(data_b, ignore_index=True)
elif len(data_b) == 1:
    b_df = data_b[0]
else:
    raise FileNotFoundError("目录B中没有找到任何符合条件的Excel文件")

# 获取目录A中的所有Excel文件
files_a = [f for f in os.listdir(dir_a) if f.endswith('.xlsx') or f.endswith('.xls')]

# 遍历目录A中的所有Excel文件
logging.info("开始处理目录A中的Excel文件...")
for file in files_a:
    filepath_a = os.path.join(dir_a, file)
    try:
        # 读取Excel文件
        df_a = pd.read_excel(filepath_a)
        
        # 将稽核资源ID转换为文本格式
        df_a['稽核资源ID'] = df_a['稽核资源ID'].astype(str)
        
        # 将目录B的数据合并到目录A的Excel文件中
        merged_df = pd.merge(df_a, b_df, left_on='稽核资源ID', right_on='eqp_id', how='left')
        
        # 打印未匹配的行信息
        unmatched_rows = df_a[~df_a['稽核资源ID'].isin(merged_df['稽核资源ID'])]
        if not unmatched_rows.empty:
            logging.warning(f"文件 {file} 中未匹配到数据的行:")
            logging.warning(unmatched_rows)
        
        # 重命名合并后的列
        merged_df.rename(columns={
            'create_date': '资源创建时间',
            'region_name': '地市'
        }, inplace=True)

        # 删除不需要的列
        merged_df.drop(columns=['eqp_id'], inplace=True, errors='ignore')

        # 重新排列列的顺序
        column_order = [
            '省份', '地市', '分类', '网络类型', '稽核规则id','稽核规则名称', '稽核资源ID', '稽核资源名称',
            '资源创建时间', '稽核失败时间', '稽核失败日期','失败原因', '稽核失败天数'
        ]

        merged_df = merged_df[column_order]

        # 保存修改后的Excel文件
        output_path = os.path.join(dir_a, f'updated_{file}')
        merged_df.to_excel(output_path, index=False)
        logging.info(f"已保存处理后的文件至 {output_path}")
    except Exception as e:
        logging.error(f"处理文件 {filepath_a} 时发生错误: {e}")

logging.info('处理完成')