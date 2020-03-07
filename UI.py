# encoding:utf-8
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import threading
import time
import Main1


def select_main_table():
    global main_table_file
    main_table_file = filedialog.askopenfilename(filetypes=[('excel file(.*xlsx)', '.xlsx')])  # 获得选择好的文件
    main_table_file = main_table_file.replace('/', '\\')
    print(main_table_file)
    main_table_text.insert('insert','主表是：' + main_table_file)
    select_main_table_button['text'] = '重新选择主表'
    print(main_table_file)
    print(len(main_table_file))
    return main_table_file


def select_sub_table():
    global sub_table_file_list
    sub_table_file_list = filedialog.askopenfilenames(filetypes=[('excel file (.*xlsx)', '.xlsx')])
    namelist = []
    str = ''
    for name in sub_table_file_list:
        name = name.replace('/', '\\')
        namelist.append(name)
        str = str + name + '\n'
    sub_table_file_list = namelist
    # 清空原来的内容
    sub_table_text.delete('1.0', 'end')
    sub_table_text.insert('insert', '子表: \n' + str)
    sub_table_text.update()
    #t['height'] = min(len(sub_table_file_list) + 1, 15)
    select_sub_table_button['text'] = '重新选择子表'
    return namelist


def call_main_begin():
    global main_table_file, sub_table_file_list, window,begin_meger_button
    begin_meger_button.place_forget()
    time.sleep(2)
    label2 = tk.Label(window, text='正在合并')
    label2.place(x=300,y=380)
    time.sleep(2)
    label2.place_forget()
    Main1.begin(main_table_file, sub_table_file_list)
    label2.config(text='合并完成')
    time.sleep(2)
    begin_meger_button.place(x=300, y=380)


def begin():
    if(len(main_table_file)==0):
        tkinter.messagebox.showwarning(title='Hi', message='主表没选')
        return
    if(len(sub_table_file_list)==0):
        tkinter.messagebox.showwarning(title='hi',message='没选从表')
        return
    thread1 = threading.Thread(target=call_main_begin)
    thread1.start()

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

#选择主表的按钮及其左边对应的text
select_main_table_button = tk.Button(content_frame,height=1,text=' 选择主表 ',background='GhostWhite',command=select_main_table)
select_main_table_button.place(x=600,y=47)
text_width = 55
main_table_text = tk.Text(content_frame,height=1,width=text_width,font=('Arial',12))
main_table_text_pos_x = 70
main_table_text.place(x=main_table_text_pos_x,y=50)

#子表的text框
sub_table_text = tk.Text(content_frame,height=14,width=text_width,font=('Arial',12))
sub_table_text.place(x=main_table_text_pos_x,y=100)
#选择子表的按钮
select_sub_table_button = tk.Button(content_frame,height=1,text=' 选择子表 ',background='GhostWhite',command=select_sub_table)
select_sub_table_button.place(x=600,y=130)
#删除子表的按钮
delete_sub_table_button = tk.Button(content_frame,height=1,text=' 删除子表 ',background='GhostWhite')
delete_sub_table_button.place(x=600,y=180)
#选择文件夹的按钮
select_folder_button = tk.Button(content_frame,height=1,text='选择文件夹',background='GhostWhite')
select_folder_button.place(x=600,y=230)
#开始合并的按钮
begin_meger_button = tk.Button(content_frame,height=1,text='开始合并',background='GhostWhite',command=begin)
begin_meger_button.place(x=300,y=380)

main_table_file = ''
sub_table_file_list = ''

tk.mainloop()
