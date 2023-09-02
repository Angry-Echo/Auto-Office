from win32com.client import Dispatch
import win32com
import win32com.client
import os

# document = Document(file_mode)
# # 读取word中的所有表格
# tables = document.tables
# document.tables[1].add_row()
app = win32com.client.Dispatch('Word.Application')
# 打开word，经测试要是绝对路径
doc = app.Documents.Open(r'E:\AngryEcho\NewLife\Script\02-开放性课题-专家签字表.doc')
# 复制word的所有内容
doc.Content.Copy()
# 关闭word
doc.Close()

word = win32com.client.DispatchEx('Word.Application')
# 创建一个新的Word文档
doc = word.Documents.Add()

# 保存新的Word文档
doc.SaveAs(r'E:\AngryEcho\NewLife\Script\new.doc')

# 关闭新的Word文档
doc.Close()

doc = word.Documents.Open(r'E:\AngryEcho\NewLife\Script\new.doc')
# myRange = doc1.Range(doc1.Content.End-1, doc1.Content.End-1)

# doc1.Range().Select()
#
# doc.myRange.Selection.Paste()
s = word.Selection
s.MoveRight(1, doc.Content.End)  # 将光标移动到文末，就这一步试了我两个多小时
s.Paste()
doc.Close()
