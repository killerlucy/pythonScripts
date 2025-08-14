import pandas as pd  
from datetime import datetime  
  
# 读取Excel文件  
file_path = 'D:\\智网创新\\故障派单失败9月\\test.xlsx'  # 输入文件的名称  
# 获取当前时间并格式化  
now = datetime.now()  
time_str = now.strftime("%Y%m%d_%H%M%S")  
# 生成Excel文件名称  
output_file_path = 'D:\\智网创新\\故障派单失败9月\\统计省份分类数量' + time_str + '.xlsx'  # 输出文件的名称  
  
# 读取数据  
df = pd.read_excel(file_path)  
  
# 定义所有可能的核查结果列  
columns_to_pivot = [  
    '【1资源类】网管割接', '【2资源类】新建资源未及时入库', '【3数据采集】采集入库延迟',  
    '【4资源类】省分维护机房信息延迟', '【5资源类】区县信息维护延迟', '【6资源类】试运行设备',  
    '【7资源类】应基站设备', '【8资源类】临时调测/测试设备', '【9资源类】设备退网/删除',  
    '【10资源类】省分未及时维护', '【11资源类或告警类】告警标识与资源标识不一致',  
    '【12非资源类】告警资源规格错误', '【13非资源类】告警关联异常', '【14非资源类】非区县级设备',  
    '【15非资源类】上报范围不一致', '【16非资源类】管理区域同步异常', '【17非资源类】无归属 或未排查',  
    '【18资源类或告警类】特殊情况', '【19数据采集类】采集数据缺失'  
]  
  
# 使用pivot_table来统计每个省份对应的核查结果数量  
# 首先，我们需要将省份核查结果列拆分成多个布尔列，表示是否包含某个特定的分类  
for col in columns_to_pivot:  
    df[col] = df['省份核查结果'].str.contains(col, na=False).astype(int)  
  
# 然后，我们使用pivot_table来汇总这些布尔列  
result_df = df.pivot_table(index='省份', values=columns_to_pivot, aggfunc='sum', fill_value=0)  
  
# 将结果保存到新的Excel文件  
result_df.reset_index().to_excel(output_file_path, index=False)  
  
print("处理完成，结果已保存到", output_file_path)