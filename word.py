import win32com.client as win32
from docx2pdf import convert

# 创建Microsoft Word应用程序对象
# 注意：一定要取消WPS文件关联和默认打开方式！！！再开启word为默认打开方式，否则无法调用该包
# 如果是非正版（未激活），记得在修改时把打开word的激活向导提示关闭
word = win32.gencache.EnsureDispatch('Word.Application')

# 打开现有的Word文档
doc = word.Documents.Open(r'E:\AngryEcho\NewLife\Script\test_word2\02-开放性课题-专家签字表.doc')

# 替换第一页的第一行文本
for paragraph in doc.Paragraphs:
    if '开放性课题' in paragraph.Range.Text:
        paragraph.Range.Text = '某项目'
        break

# 复制第一页并创建新页
for i in range(1, 10):  # 假设需要创建10页
    section = doc.Sections(1)  # 获取第一页的页面设置
    new_page = section.Range.Copy()  # 复制第一页
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(7)  # 在文档末尾插入分页符
    doc.Range(doc.Content.End - 1, doc.Content.End - 1).Paste()  # 粘贴复制的内容

    # # 替换新页面的第一行文本
    # for paragraph in new_page.Paragraphs:
    #     if '开放性课题' in paragraph.Range.Text:
    #         paragraph.Range.Text = '新的文本'
    #         break

# 将整个文档保存为PDF
doc.SaveAs('example.docx', FileFormat=16)  # 将文档另存为.docx格式
doc.Close()  # 关闭文档
word.Quit()  # 关闭Word应用程序
# convert('example.docx') # 转换为PDF

# 重新运行时一定要关闭打开的文件！！！
# 复制了单元格样式
# 要绝对路径！！！
