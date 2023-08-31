import pandas as pd
import os

root_path = r'E:\AngryEcho\NewLife\Script\test_dir'

# 获取某个文件夹中的所有'.xlsx'后缀的文件名
file_names = [filename for filename in os.listdir(root_path) if filename.endswith('.xlsx')]

# 读取所有Excel文件中的表并合并为一个DataFrame
dfs = []
for file_name in file_names:
    file_path = os.path.join(root_path, file_name)
    workbook = pd.read_excel(file_path, sheet_name='Sheet1')
    dfs.append(workbook)

merged = pd.concat(dfs, axis=0, ignore_index=False)

merged.to_excel('./合并表1.xlsx', index=False)