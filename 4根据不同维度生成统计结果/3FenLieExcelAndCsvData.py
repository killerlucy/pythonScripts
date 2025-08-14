import os
import pandas as pd

# 定义源目录和目标目录
source_dir = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\分裂前的数据'
target_dir = r'D:\智网创新\无线网资源运营月报\月报ppt\11月\11月省份稽核规则下载\分裂后的数据'

# 检查目标目录是否存在，如果不存在则创建
if not os.path.exists(target_dir):
    os.makedirs(target_dir)

# 遍历源目录中的所有文件
for filename in os.listdir(source_dir):
    if filename.endswith('.txt'):
        # 构建完整的文件路径
        file_path = os.path.join(source_dir, filename)
        
        # 读取文件内容
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        # 分割每一行的内容
        data = [line.strip().split('￥') for line in lines]
        
        # 将数据转换为DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])
        
        # 构建目标文件路径
        target_file_path = os.path.join(target_dir, os.path.splitext(filename)[0] + '.xlsx')
        
        # 写入Excel文件
        df.to_excel(target_file_path, index=False, engine='openpyxl')

print("处理完成")