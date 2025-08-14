import os  
import subprocess  
  
directory = 'D:/智网创新/python程序/2查找降低的省份'  # 替换为你的目录路径  
  
for filename in os.listdir(directory):  
    if filename.endswith('.py') and filename != '__main__.py':  # 排除当前脚本（如果有的话）  
        file_path = os.path.join(directory, filename)  
        try:  
            result = subprocess.run(['python', file_path], capture_output=True, text=True, check=True)  
            print(f'Executed {filename}:\n{result.stdout}')  
        except subprocess.CalledProcessError as e:  
            print(f'Error executing {filename}:\n{e.stderr}')