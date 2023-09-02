import openpyxl
import glob
import os

# 修改目录work_dir
work_dir = r'./test_dir'

# 遍历所有excel文件的sheet,存为list
sheet_list = []
for path in glob.glob(os.path.join(work_dir, '*.xlsx')):  # 搜索给定目录下所有的.xlsx文件
    wb = openpyxl.load_workbook(path)
    sheet = wb.get_sheet_by_name('Sheet1')
    sheet_list.append(sheet)


# print(type(ws.cell(1,2).value))


def get_max_row(sheet):
    real_max_row = 1

    while True:
        row_dict = {cell.value for cell in sheet[real_max_row]}
        if row_dict != {None}:
            real_max_row += 1
        else:
            break

    return real_max_row - 1


j = get_max_row(sheet_list[0])
print("通过自定义函数获取到的最大行是：", j)
