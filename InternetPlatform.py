import pandas as pd
import os

# 获取某个文件夹中的所有'.xlsx'后缀的文件名
excel_files = [filename for filename in os.listdir() if filename.endswith('.xlsx')]

# 读取所有Excel文件中的表并合并为一个DataFrame
dfs = []
for filename in excel_files:
    workbook = pd.read_excel(filename, sheet_name=None)
    first_sheet_name = list(workbook.keys())[0]  # 获取第一个工作表的名称
    first_sheet = workbook[first_sheet_name]  # 获取第一个工作表
    if len(dfs) > 0:
        # 从第二行开始读取第一个工作表的数据
        first_sheet = first_sheet.iloc[1:]
    else:
        # 只保留第一个工作表的第一行
        first_sheet = first_sheet.iloc[[0]]
    first_sheet['文件名'] = filename  # 在表中添加一列文件名，用于区分不同的Excel文件
    first_sheet['工作表名'] = first_sheet_name  # 在表中添加一列工作表名，用于区分不同的工作表
    dfs.append(first_sheet)
    for sheet_name, df in workbook.items():
        if sheet_name != first_sheet_name:
            # 从第二行开始读取其他工作表的数据
            df = df.iloc[1:]
            df['文件名'] = filename  # 在表中添加一列文件名，用于区分不同的Excel文件
            df['工作表名'] = sheet_name  # 在表中添加一列工作表名，用于区分不同的工作表
            dfs.append(df)
merged_df = pd.concat(dfs, ignore_index=True)

# 将合并后的DataFrame保存为Excel文件
merged_df.to_excel('merged.xlsx', index=False)