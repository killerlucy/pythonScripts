import os
import pandas as pd
import logging

# 设置日志配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义目录路径
base_dir = r'D:\智网创新\无线网资源运营月报'
analysis_file = os.path.join(base_dir, '月报ppt', '10月', '10月稽核规则统计汇总', '10月分析稽核规则.xlsx')
detail_file = os.path.join(base_dir, '月报ppt', '10月', '10月稽核规则统计汇总', '10月稽核明细分析汇总.xlsx')
auto_collect_file = os.path.join(base_dir, '自动采集.xlsx')

# 读取“10月分析稽核规则”文件
try:
    analysis_df = pd.read_excel(analysis_file, usecols=[
        '分类', '网络类型', '稽核规则名称', '省份', '导致省分稽核成功率低', '导致省分稽核成功率下降'
    ])
    logging.info("成功读取10月分析稽核规则文件")
except Exception as e:
    logging.error(f"无法读取10月分析稽核规则文件: {e}")
    raise

# 向表A中追加列
analysis_df['存量资源'] = None
analysis_df['新增资源'] = None
analysis_df['总计'] = None
analysis_df['维护方式'] = None
analysis_df['原因'] = None
analysis_df['举措'] = None
analysis_df['时限'] = None
logging.info("表A中已追加新列")

# 读取“10月稽核明细分析汇总”文件
try:
    detail_df = pd.read_excel(detail_file, usecols=[
        '分类', '网络类型', '省份', '稽核规则名称', '存量资源', '新增资源', '总计'
    ])
    logging.info("成功读取10月稽核明细分析汇总文件")
except Exception as e:
    logging.error(f"无法读取10月稽核明细分析汇总文件: {e}")
    raise

# 以内容A为主表，根据内容A和内容B中的“分类    网络类型    稽核规则名称    省份”四列进行匹配
merge_keys = ['分类', '网络类型', '省份', '稽核规则名称']
analysis_df = pd.merge(analysis_df, detail_df, on=merge_keys, how='left')
logging.info("匹配并填充存量资源、新增资源、总计列")

# 读取“自动采集”文件
try:
    auto_collect_df = pd.read_excel(auto_collect_file, usecols=['稽核规则名称'])
    logging.info("成功读取自动采集文件")
except Exception as e:
    logging.error(f"无法读取自动采集文件: {e}")
    raise

# 根据“10月分析稽核规则”文件中的“稽核规则名称”列的内容，匹配“自动采集”文件中“稽核规则名称”列中的信息
auto_rules = set(auto_collect_df['稽核规则名称'])
analysis_df['维护方式'] = analysis_df['稽核规则名称'].apply(lambda x: '自动采集' if x in auto_rules else '手动维护')
logging.info("维护方式列已更新")

# 保存更新后的文件
output_file = os.path.join(base_dir, '月报ppt', '10月', '10月稽核规则统计汇总', '更新后的10月分析稽核规则.xlsx')
analysis_df.to_excel(output_file, index=False)
logging.info(f"已保存更新后的文件至 {output_file}")

logging.info('处理完成')