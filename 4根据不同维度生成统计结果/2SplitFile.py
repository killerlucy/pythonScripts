import os  
import pandas as pd  
  
def split_excel_files(input_dir, output_dir):  
    # 确保输出目录存在  
    if not os.path.exists(output_dir):  
        os.makedirs(output_dir)  
      
    # 获取目录下所有Excel文件  
    excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') or f.endswith('.xls')]  
      
    for file_name in excel_files:  
        file_path = os.path.join(input_dir, file_name)  
          
        # 读取Excel文件  
        df = pd.read_excel(file_path, engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd')  
          
        # 计算需要拆分的文件数量  
        rows_per_file = 3000  
        num_files = (len(df) // rows_per_file) + (1 if len(df) % rows_per_file != 0 else 0)  
          
        base_name = os.path.splitext(file_name)[0]  
        extension = os.path.splitext(file_name)[1]  
          
        # 拆分并保存文件  
        for i in range(num_files):  
            start_row = i * rows_per_file  
            end_row = min((i + 1) * rows_per_file, len(df))  
            split_df = df.iloc[start_row:end_row]  
              
            split_file_name = f"{base_name}{i + 1}{extension}"  
            split_file_path = os.path.join(output_dir, split_file_name)  
              
            split_df.to_excel(split_file_path, index=False, engine='openpyxl' if extension == '.xlsx' else 'xlwt')  
            print(f"Saved {split_file_name} to {output_dir}")  
  
# 调用函数，指定输入和输出目录  
input_directory = 'D:\\智网创新\\无线网资源运营月报\\月报ppt\\9月\\9月稽核明细\\数据量大需要拆分的文件'  # 替换为你的输入目录路径  
output_directory = 'D:\\智网创新\\无线网资源运营月报\\月报ppt\\9月\\9月稽核明细\\数据量大需要拆分的文件结果'  # 替换为你的输出目录路径  
split_excel_files(input_directory, output_directory)