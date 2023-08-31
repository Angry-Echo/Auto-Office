from win32com.client import Dispatch
import openpyxl

# 打开Excel文件并读取数据
wb = openpyxl.load_workbook('example.xlsx')
ws = wb.active
data = [cell.value for cell in ws['A'][1:]]  # 假设需要替换的文本在第一列


app = Dispatch('Word.Application')
app.Visible = True
doc = app.Documents.Open(r'E:\AngryEcho\NewLife\Script\02-开放性课题-专家签字表.doc')
# 复制word的所有内容
doc.Content.Copy()
# 关闭word
doc.Close()

s = app.Selection
s.Find.Execute('开放性课题', False, False, False, False, False, True, 1, False, '某某', 2)

# 复制第一页并创建新页
for i in range(1, 3):  # 假设需要创建10页  指数级增长
    section = doc.Sections(1)  # 获取第一页的页面设置
    new_page = section.Range.Copy()  # 复制第一页
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(7)  # 在文档末尾插入分页符
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).Paste()  # 粘贴复制的内容
#
#     # 替换新页面的第一行文本
#     for paragraph in new_page.Paragraphs:
#         if '需要替换的文本' in paragraph.Range.Text:
#             paragraph.Range.Text = '新的文本'
#             break
# 将整个文档保存为PDF
# doc.SaveAs(r'E:\AngryEcho\NewLife\Script\test_word2\02-开放性课题-专家签字表.doc', FileFormat=16) # 将文档另存为.docx格式
# doc.Close() # 关闭文档
# word.Quit() # 关闭Word应用程序
# convert('example.docx') # 转换为PDF
