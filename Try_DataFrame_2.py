"""
Created on Wednesday, March 25, 2020 at 11:14:56

@author: qinghua mao
"""

import os, time
import pandas as pd

start_time = time.time()
dir = r'D:\python脚本\合并excel\日常指标'  # 设置工作路径

# 新建列表，存放每个文件数据框（每一个excel读取后存放在数据框,依次读取多个相同结构的Excel文件并创建DataFrame）
DFs = []

for root, dirs, files in os.walk(dir):  # 第一个为起始路径，第二个为起始路径下的文件夹，第三个是起始路径下的文件。
    for file in files:
        file_path = os.path.join(root, file)  # 将路径名和文件名组合成一个完整路径
        df = pd.read_excel(file_path, encoding="gbk")  # excel转换成DataFrame
        DFs.append(df)
# 合并所有数据，将多个DataFrame合并为一个
alldata = pd.concat(DFs)  # sort='False'

# alldata.to_csv(r'D:\python脚本\csv合并结果.csv',sep=',',index = False,encoding="gbk")
alldata.to_excel(r'D:\python脚本\excel合并结果.xlsx', index=False, encoding="gbk")
end_time = time.time()
times = round(end_time - start_time, 2)
print('合并完成，耗时{}秒'.format(times))
# 如果要将合并结果写入到csv文件中，就使用 to_csv,如果要将合并结果写入到excel文件中，就使用 to_excel
# 如果是合并带有数字的excel，最好写入到csv文件中（个人建议），写入到excel中还需要将数字单元格进行转换,但是如果有日期，需要手动转换
# 如果写入结果，中文有乱码，就指定写入格式，这里指定的是gbk
