import threading
from time import sleep
from tkinter import *


# class Main():
#     def __init__(self):
#         root = Tk()
#         self.entry = Entry(root)
#         self.button = Button(root, text='执行', command=self.do_something)
#         self.grid()
#         root.mainloop()
#
#     def grid(self):
#         self.entry.grid(row=0, column=0)
#         self.button.grid(row=0, column=1)
#
#     def do_something(self):
#         time.sleep(10)
#         self.entry.insert('0', '结束了')


class Main():
    def __init__(self):
        root = Tk()
        self.entry = Entry(root)
        time = 10
        self.button = Button(root, text='执行',
                             command=lambda: threading.Thread(
                                 target=self.do_something,
                                 args=(time,)
                             ).start())
        self.grid()
        root.mainloop()

    def grid(self):
        self.entry.grid(row=0, column=0)
        self.button.grid(row=0, column=1)

    def do_something(self, data):
        for i in range(data, 0, -1):
            self.entry.insert('0', f'倒计时{i}秒')
            sleep(1)
            self.entry.delete('0', END)
        else:
            self.entry.insert('0', '倒计时结束')


if __name__ == '__main__':
    Main()
