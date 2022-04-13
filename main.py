import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox as msg
import winreg


def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]


def execute():
    f = open(get_desktop()+os.sep+'result.vcf','w',encoding='utf-8')
    text = edit.get(0.0,tk.END)
    lines = str(text).split('\n')
    for line in lines:
        if line:
            ori_data = line.split()
            if len(ori_data)>2 or ori_data[0].isdigit() or not ori_data[1].isdigit():
                msg.showerror('错误', '请检查格式，一行只能有姓名和电话。')
                return
            new_data = f'BEGIN:VCARD\nVERSION:3.0\nFN:{ori_data[0]}\nTEL;VALUE=text:{ori_data[1]}\nEND:VCARD\n'
            f.write(new_data)
    f.close()
    msg.showinfo('提示','成功！已保存到桌面！')


mainForm = tk.Tk()
mainForm.title('一键生成通讯录')

edit = tk.Text(mainForm, height=25, width=35)
edit.insert(0.0,"地鼠\t18012345678\n张三\t13012345678\n↑按照这个格式把名单复制进来\n（从Excel文档直接复制就行）\nAuthor:地鼠")
edit.pack(padx=5,pady=5)

button = ttk.Button(mainForm, text="执行", width=35, command=execute)
button.pack(padx=5,pady=5)

mainForm.mainloop()
