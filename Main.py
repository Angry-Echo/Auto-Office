from Merge_2 import Merge_Excel
from Word import Word_PDF

root_dir = input('请输入您的很多个Excel表所在的目录（绝对路径）:')
example_path = input('请输入示例的专家签字表的路径：')

Merge_Excel(root_dir)
Word_PDF(root_dir, example_path)
