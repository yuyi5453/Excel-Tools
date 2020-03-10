# encoding:utf-8
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import threading
import time
import Main1
import sys
import re
from Enable_write import *
import judge_num
#关闭窗口时间
def close_win():
    window.destroy()
    return

def select_main_table():
    global main_table_file
    main_table_file = filedialog.askopenfilename(filetypes=[('excel file(.*xlsx)', '.xlsx')])  # 获得选择好的文件
    main_table_file = main_table_file.replace('/', '\\')
    if is_only_read(main_table_file):
        tkinter.messagebox.showwarning(title='error', message='主表文件已被某一应用程序打开，请关闭后重试')
        return
    main_table_text.delete(1.0,'end')
    main_table_text.insert('insert', '主表是：' + main_table_file)
    select_main_table_button['text'] = '重新选择主表'
    return main_table_file


def select_sub_table():
    global sub_table_file_list
    sub_table_file_list = filedialog.askopenfilenames(filetypes=[('excel file (.*xlsx)', '.xlsx')])
    namelist = []
    #清空子表
    sub_table_list.delete(0,'end')
    for name in sub_table_file_list:
        name = name.replace('/', '\\')
        namelist.append(name)
        #将子表加入listbox中
        sub_table_list.insert('end',name)
    # sub_table_file_list = namelist
    #t['height'] = min(len(sub_table_file_list) + 1, 15)
    select_sub_table_button['text'] = '重新选择子表'
    return namelist


def call_main_begin(key,need_add):
    global main_table_file, sub_table_file_list, window,begin_meger_button
    begin_meger_button.place_forget()
    time.sleep(2)
    label2 = tk.Label(window, text='正在合并')
    label2.place(x=300,y=380)
    time.sleep(2)
    label2.place_forget()
    Main1.begin(main_table_file, sub_table_file_list,key,need_add)
    label2.config(text='合并完成')
    time.sleep(2)
    begin_meger_button.place(x=300, y=380)
    sub_table_list.config(state=tk.NORMAL)

def begin():
    global sub_table_file_list, input_key_entry,input_addble_col_entry
    key = input_key_entry.get()
    print(type(key),key)
    if(key==''):
        tkinter.messagebox.showwarning(title='error', message='请输入主键')
        return
    if(judge_num.is_integer(key)==False):
        tkinter.messagebox.showwarning(title='error', message='主键格式应为数字')
        return
    key = int(key)
    need_add = []
    addble_col = input_addble_col_entry.get()
    need_add_list = [int(s) for s in re.findall(r'\b\d+\b', addble_col)]
    need_add = tuple(need_add_list)
    print(need_add)
    if(len(main_table_file) == 0):
        tkinter.messagebox.showwarning(title='error', message='主表没选')
        return

    if is_only_read(main_table_file):
        tkinter.messagebox.showwarning(title='error', message='主表文件已被某一应用程序打开，请关闭后重试')
        return
    sub_table_list.config(state=tk.DISABLED)

    namelist=[]
    for i in range(0,sub_table_list.size()):
        namelist.append(sub_table_list.get(i))
        if is_only_read(sub_table_list.get(i)):
            tkinter.messagebox.showwarning(title='error', message='子表文件'+sub_table_list.get(i)+'已被某一应用程序打开，请关闭后重试')
            return
    sub_table_file_list=namelist
    if(len(sub_table_file_list)==0):
        tkinter.messagebox.showwarning(title='hi',message='没选从表')
        return
    thread1 = threading.Thread(target=call_main_begin,args=(key,need_add))
    thread1.daemon=True
    thread1.start()

#删除选择的子表
def delete_sub_table(event):
    pos=sub_table_list.curselection()
    sub_table_list.delete(pos)
    return

#创建主窗口
window = tk.Tk()
window.title('Excel tool')
window.geometry('800x500')

#创建主frame
frame = tk.Frame(window,width=800,height=500)
frame.pack()

#navigation_frame: 左边的导航界面
navigation_frame = tk.Frame(frame,width=100,height=500,bg='green')
content_frame = tk.Frame(frame,width=700,height=500)
navigation_frame.place(x=0,y=0)
content_frame.place(x=100,y=0)


#布置导航界面的图片
canvas = tk.Canvas(navigation_frame, bg='blue', height=500, width=100)
canvas.place(x=0,y=0)
image_file = tk.PhotoImage(file='img\p1.jpg')
image = canvas.create_image(0, 0, anchor='nw', image=image_file)

#导航界面的lookup按钮和合并按钮
lookup_button = tk.Button(navigation_frame,height=2,text='lookup\n函数生成',background='GhostWhite')
lookup_button.place(x=20,y=130)
meger_button = tk.Button(navigation_frame,height=2,text='合并表格',background='GhostWhite')
meger_button.place(x=20,y=260)


dealt_y = 30

label = tk.Label(content_frame,text='主健：')
label.place(x=70, y=20)
input_key_entry = tk.Entry(content_frame,width=12,font=('Arial',12))
input_key_entry.place(x=70, y=40)

label = tk.Label(content_frame,text='可加列：')
label.place(x=270, y=20)
input_addble_col_entry = tk.Entry(content_frame,width=12,font=('Arial',12))
input_addble_col_entry.place(x=270, y=40)



#选择主表的按钮及其左边对应的text
select_main_table_button = tk.Button(content_frame,height=1,text=' 选择主表 ',background='GhostWhite',command=select_main_table)
select_main_table_button.place(x=600,y=47+dealt_y)
text_width = 55
main_table_text = tk.Text(content_frame,height=1,width=text_width,font=('Arial',12))
main_table_text_pos_x = 70
main_table_text.place(x=main_table_text_pos_x,y=50+dealt_y)

#子表的listbox框
sub_table_list=tk.Listbox(content_frame,height=14,width=text_width,font=('Arial',12))
sub_table_list.place(x=main_table_text_pos_x,y=100+dealt_y)
sub_table_list.bind('<Double-Button-1>',delete_sub_table)



#提示
lab=tk.Label(content_frame,text='*双击删除',fg='red')
lab.place(x=main_table_text_pos_x, y=380+dealt_y)
#选择子表的按钮
select_sub_table_button = tk.Button(content_frame,height=1,text=' 选择子表 ',background='GhostWhite',command=select_sub_table)
select_sub_table_button.place(x=600, y=130+dealt_y)
#选择文件夹的按钮
select_folder_button = tk.Button(content_frame,height=1,text='选择文件夹',background='GhostWhite')
select_folder_button.place(x=600, y=230+dealt_y)
#开始合并的按钮
begin_meger_button = tk.Button(content_frame,height=1,text='开始合并',background='GhostWhite',command=begin)
begin_meger_button.place(x=300,y=400+dealt_y)

main_table_file = ''
sub_table_file_list = ''
window.protocol('WM_DELETE_WINDOW',close_win)
tk.mainloop()
