import merge_excel
from datetime import datetime

# 定义参数
root_dir = 'D:/智网创新/2月报/月报ppt/1月/第一步下载稽核数据/'
target_dir = '跨网络/跨网络4G/'
output_file_name = '跨网络4G'
file_order = [  
    "BBU关联机房放置点关联稽核",  
    "4GRRU关联机房放置点关联稽核",  
    "EUTRANCELL关联RRU所属机房经纬度完整性稽核(不含中兴NB小区)"
]

# 拼接路径
directory = root_dir + target_dir
output_file = root_dir + output_file_name + '_' + datetime.now().strftime("%Y%m%d_%H%M%S") + '.xlsx'
# 调用函数
merge_excel.merge_excel_files(directory, file_order, output_file)