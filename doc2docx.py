import os
from win32com import client as wc


def TransDocToDocx(oldDocName, newDocxName):
    print("我是 TransDocToDocx 函数")
    # 打开word应用程序
    word = wc.Dispatch('Word.Application')

    # 打开 旧word 文件
    doc = word.Documents.Open(oldDocName)

    # 保存为 新word 文件,其中参数 12 表示的是docx文件
    doc.SaveAs(newDocxName, 12)

    # 关闭word文档
    doc.Close()
    word.Quit()

    print("生成完毕！")


if __name__ == "__main__":
    # 获取当前目录完整路径
    currentPath = os.getcwd()
    print("当前路径为：", currentPath)

    # 获取 旧doc格式word文件绝对路径名
    docName = os.path.join(currentPath, '02-开放性课题-专家签字表.doc')
    print("docFilePath = ", docName)

    # 设置新docx格式文档文件名
    docxName = os.path.join(currentPath, 'test.docx')

    TransDocToDocx(docName, docxName)
