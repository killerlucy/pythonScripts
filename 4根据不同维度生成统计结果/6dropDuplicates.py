import os
import pandas as pd

# 文件夹路径
directory = 'D:/智网创新/无线网资源运营月报/月报ppt/9月/9月稽核明细统计结果4类/3result/统计结果汇总/drop'  # 替换为你的Excel文件夹路径

# 遍历文件夹中的所有 Excel 文件
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        file_path = os.path.join(directory, filename)
        
        # 读取 Excel 文件
        df = pd.read_excel(file_path)
        
        # 对当前文件进行去重处理
        unique_df = df.drop_duplicates(subset=['稽核资源ID', '稽核资源名称'], keep='first')

        # 新文件名，可以加上 "_unique" 或其他标识
        new_filename = f"{os.path.splitext(filename)[0]}_unique.xlsx"
        new_file_path = os.path.join(directory, new_filename)
        
        # 保存去重后的数据到新的 Excel 文件
        unique_df.to_excel(new_file_path, index=False)

print(f"所有文件已完成去重处理并保存在 {directory} 目录下")