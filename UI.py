#coding=gbk
import tkinter as tk
import xlwings as xw
from tkinter import filedialog
import threading
import Main
import time
from PIL import Image

from PIL import ImageTk
def select_main_table():
    global main_tabel_file
    main_tabel_file = filedialog.askopenfilename(filetypes=[('excel file(.*xlsx)','.xlsx')]) #获得选择好的文件
    main_tabel_file = main_tabel_file.replace('/','\\')
    L1.config(text='主表是：'+main_tabel_file)
    button1['text'] = '重新选择主表'
    return main_tabel_file

def select_sub_table():
    global  sub_table_file_list
    sub_table_file_list = filedialog.askopenfilenames(filetypes=[('excel file (.*xlsx)','.xlsx')])
    namelist = []
    str = ''
    for name in sub_table_file_list:
        name = name.replace('/','\\')
        namelist.append(name)
        str = str + name + '\n'
    sub_table_file_list = namelist
    #清空原来的内容
    t.delete('1.0','end')
    t.insert('insert','子表: \n' + str)
    t.update()
    t['height'] = min(len(sub_table_file_list)+1,15)
    button2['text'] = '重新选择子表'
    return namelist
def call_main_begin():
    global main_tabel_file , sub_table_file_list  ,window
    button3.pack_forget()
    label2 = tk.Label(window, text='正在合并')
    label2.pack()
    Main.begin(main_tabel_file,sub_table_file_list)
    label2.config(text='合并完成')
    button3.pack()
def begin():
    thread1 = threading.Thread(target=call_main_begin)
    thread1.start()

        #初始化窗口大小
window = tk.Tk()
window.title('this is a window')
window.geometry('700x700')

main_tabel_file = ''
sub_table_file_list = []
        #选择主表的按钮以及标签


L1 = tk.Label(window,text='请选择主表',font=('Arial',14))
L1.pack()
button1 = tk.Button(window,text='选择主表',command = select_main_table,bg='#F0FFFF')
button1.pack()

        #选择子表的按钮及文本
t = tk.Text(window,height = 5,font=('Arial',14))
t.insert('insert','子表:\n')
t.pack()
button2 = tk.Button(window,text='选择子表',command=select_sub_table,bg='#F0FFFF')
button2.pack()

button3 = tk.Button(window,text='开始合并',command = begin)
button3.pack()


window.mainloop()