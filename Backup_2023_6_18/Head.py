import openpyxl
import glob
import os
import sys


def Find_Max_Head(sheet_list):

    row_flag = 1
    try:

        wb_1_sheet = sheet_list[0]
        wb_2_sheet = sheet_list[1]

        while True:
            wb_1_sheet_row = {cell.value for cell in wb_1_sheet[row_flag]}
            wb_2_sheet_row = {cell.value for cell in wb_2_sheet[row_flag]}

            if wb_1_sheet_row == wb_2_sheet_row:
                row_flag += 1
            else:
                break

    except PermissionError:
        print('有Excel文件未关闭！请关闭后重试')
        sys.exit()

    return row_flag - 1  # row_flag比最大行数多一个，是非表头行的位置

