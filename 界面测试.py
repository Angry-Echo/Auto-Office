import tkinter
from tkinter import filedialog
from tkinter import Spinbox
from tkinter import Checkbutton

Window = tkinter.Tk()

Window.title('请选择您想要合并的Excel表所在的目录')

# GUI界面在屏幕中间出现
max_w, max_h = Window.maxsize()
Window.geometry(f'500x500+{int((max_w - 500) / 2)}+{int((max_h - 500) / 2)}')
Window.resizable(width=False, height=False)

# 标签
label = tkinter.Label(Window, text='选择目录：', font=('宋体', 15))
label.place(x=42, y=50)

# 输入框
entry_text = tkinter.StringVar(Window)
entry = tkinter.Entry(Window, textvariable=entry_text, font=('FangSong', 10), width=35, state='readonly')
entry.place(x=150, y=55)


# 按钮控件
def get_dir_path():
    """注意，以下列出的方法都是返回字符串而不是数据流"""

    # 返回一个字符串，且只能获取文件夹路径，不能获取文件的路径。
    path = filedialog.askdirectory(title='请选择一个目录')

    entry_text.set(path)


# def get_file_path():

# 返回一个字符串，可以获取到任意文件的路径。
# path = filedialog.askopenfilename(title='请选择文件')

# 生成保存文件的对话框， 选择的是一个文件而不是一个文件夹，返回一个字符串
# path = filedialog.asksaveasfilename(title='请输入保存的路径')

button = tkinter.Button(Window, text='选择路径', command=get_dir_path)
button.place(x=410, y=50)

# 标签
label = tkinter.Label(Window, text='表头所在行：', font=('宋体', 15))
label.place(x=15, y=80)

spin = Spinbox(Window, from_=0, to=100, width=10)
spin.place(x=150, y=80)

label = tkinter.Label(Window, text='序号所在列：', font=('宋体', 15))
label.place(x=250, y=80)

spin = Spinbox(Window, from_=0, to=100, width=10)
spin.place(x=350, y=80)

chk_state = tkinter.BooleanVar()
chk_state.set(True)  # Set check state
chk = Checkbutton(Window, font=('宋体', 15), text="是否修改字体（字号--14 字体--仿宋）", var=chk_state)
chk.place(x=45, y=200)

Window.mainloop()
