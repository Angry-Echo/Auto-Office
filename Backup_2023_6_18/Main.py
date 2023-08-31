from Merge_2 import Merge_Excel
from Word import Word_PDF

root_dir = input('请输入Excel表所在目录（绝对路径）:')
head_row = int(input('请告诉我表头在第几行（数字）:'))
col_id = int(input("请输入表中的'序号'列所在的列号（数字）:"))
font_flag = input('是否需要修改表中字体（字号--14 字体--仿宋）:')
Merge_Excel(root_dir, head_row, col_id, font_flag)

keep_dialog = input('是否继续生成专家名单（请输入 是 或 否）:')

if keep_dialog == '是':
    example_path = input('请输入示例的专家签字表的路径及文件名：')
    object_str = input('请输入检测的签字表标题：')
    proj_col_id = input('请输入“项目名称”列的列号（如果是第一列应该输入A）：')
    Word_PDF(root_dir, example_path, object_str, proj_col_id, head_row)
    input('请按下回车键以退出程序')

else:
    input('请按下回车键以退出程序')
# C:\Users\xxx\Documents\c.doc 开放性课题项目
