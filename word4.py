import win32com.client as win32

# 创建Word应用程序对象
word = win32.Dispatch('Word.Application')
word.Visible = True

# 打开要复制内容的Word文档
doc1 = word.Documents.Open(r'E:\AngryEcho\NewLife\Script\02-开放性课题-专家签字表.doc')

# 选择要复制的内容
selection = word.Selection
selection.WholeStory()
selection.Copy()

# 将内容粘贴到新的Word文档中
# 打开新的Word文档
doc = word.Documents.Add()
selection2 = word.Selection
selection2.EndKey(win32.constants.wdStory)
selection2.Paste()

selection2.EndKey(win32.constants.wdStory)
selection2.InsertBreak(win32.constants.wdPageBreak)

# 保存并关闭新的Word文档
# 保存新的Word文档
doc.SaveAs(r'E:\AngryEcho\NewLife\Script\4_6.docx')
doc.Close()

doc = word.Documents.Open(r'E:\AngryEcho\NewLife\Script\4_6.docx')
for i in range(5):
    selection2 = word.Selection
    selection2.EndKey(win32.constants.wdStory)
    selection2.Paste()

    selection2.EndKey(win32.constants.wdStory)
    selection2.InsertBreak(win32.constants.wdPageBreak)

doc1.Close()
# 关闭Word应用程序
word.Quit()
