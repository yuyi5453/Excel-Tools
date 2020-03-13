#coding=utf-8

import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
import threading
import time
import Main
import sys
import re
from Enable_write import *

class merger_UI:
    def close_meger_content(self):
        #self.root.destroy()
        self.content_frame.place_forget()
        return

    def select_main_table(self):
        self.main_table_file = filedialog.askopenfilename(filetypes=[('excel file(.*xlsx)', '.xlsx')])  # 获得选择好的文件
        self.main_table_file = self.main_table_file.replace('/', '\\')
        if is_only_read(self.main_table_file):
            tkinter.messagebox.showwarning(title='error', message='主表文件已被某一应用程序打开，请关闭后重试')
            return
        self.main_table_text.delete(1.0, 'end')
        self.main_table_text.insert('insert', '主表是：' + self.main_table_file)
        self.select_main_table_button['text'] = '重新选择主表'

    def select_sub_table(self):
        self.sub_table_file_list = filedialog.askopenfilenames(filetypes=[('excel file (.*xlsx)', '.xlsx')])
        namelist = []
        # 清空子表
        self.sub_table_list.delete(0, 'end')
        for name in self.sub_table_file_list:
            name = name.replace('/', '\\')
            namelist.append(name)
            # 将子表加入listbox中
            self.sub_table_list.insert('end', name)
        self.select_sub_table_button['text'] = '重新选择子表'
        return namelist

    def call_main_begin(self,key, need_add):
        self.begin_meger_button.config(state=tk.DISABLED)
        time.sleep(2)
        self.label2 = tk.Label(self.content_frame,text='Details')
        self.label2.place(x=70, y=460)
        time.sleep(2)
        Main.begin(self.main_table_file, self.sub_table_file_list, key, need_add,self.label2)
        time.sleep(2)
        self.begin_meger_button.config(state=tk.NORMAL)
        self.sub_table_list.config(state=tk.NORMAL)

    def is_digit(self,c):
        return c >= '0' and c <= '9'


    def is_integer(self,str):
        for c in str:
            if (self.is_digit(c) == False):
                return False
        return True

    def begin(self):
        key = self.input_key_entry.get()
        # print(type(key), key)
        if (key == ''):
            tkinter.messagebox.showwarning(title='error', message='请输入主键')
            return
        if (self.is_integer(key) == True):
            key = int(key)
        elif(len(key)>1):
            tkinter.messagebox.showwarning(title='error', message='输入主键不是数字时长度应小于1')
            return
        else:
            key = key.lower()
            key = ord(key) - 96
        need_add = []

        addble_col = self.input_addble_col_entry.get()


        list1 = need_add_list = [int(s) for s in re.findall(r'\b\d+\b', addble_col)]
        result = ''.join(re.findall(r'[A-Za-z]', addble_col))
        for s in result:
            num = ord(s) - 96
            if num not in list1:
                list1.append(num)
        list1.sort()
        need_add = tuple(list1)


        if (len(self.main_table_file) == 0):
            tkinter.messagebox.showwarning(title='error', message='主表没选')
            return

        if is_only_read(self.main_table_file):
            tkinter.messagebox.showwarning(title='error', message='主表文件已被某一应用程序打开，请关闭后重试')
            return
        self.sub_table_list.config(state=tk.DISABLED)

        namelist = []
        for i in range(0, self.sub_table_list.size()):
            namelist.append(self.sub_table_list.get(i))
            if is_only_read(self.sub_table_list.get(i)):
                tkinter.messagebox.showwarning(title='error',message='子表文件' + self.sub_table_list.get(i) + '已被某一应用程序打开，请关闭后重试')
                return

        self.sub_table_file_list = namelist
        if (len(self.sub_table_file_list) == 0):
            tkinter.messagebox.showwarning(title='hi', message='没选从表')
            return

        # for s in self.sub_table_file_list:
        #     print(s)
        # for x in need_add_list:
        #     print(x)
        thread1 = threading.Thread(target=self.call_main_begin, args=(key, need_add))
        thread1.daemon = True
        thread1.start()

    # 删除选择的子表
    def delete_sub_table(self,event):
        pos = self.sub_table_list.curselection()
        self.sub_table_list.delete(pos)
        return

    def __init__(self,window,frame_pos_x=100,frame_pos_y=0):
        self.content_frame = tk.Frame(window,width=700,height=500)
        #self.content_frame.place(x=100, y=0)
        self.root = window
        dealt_y = 30

        self.label = tk.Label(self.content_frame,text='主健：')
        self.label.place(x=70, y=20)
        self.input_key_entry = tk.Entry(self.content_frame,width=12,font=('Arial',12))
        self.input_key_entry.place(x=70, y=40)

        self.label = tk.Label(self.content_frame,text='可加列：')
        self.label.place(x=270, y=20)
        self.input_addble_col_entry = tk.Entry(self.content_frame,width=12,font=('Arial',12))
        self.input_addble_col_entry.place(x=270, y=40)



        #选择主表的按钮及其左边对应的text
        self.select_main_table_button = tk.Button(self.content_frame,height=1,text=' 选择主表 ',background='GhostWhite',command=self.select_main_table)
        self.select_main_table_button.place(x=600,y=47+dealt_y)
        text_width = 55
        self.main_table_text = tk.Text(self.content_frame,height=1,width=text_width,font=('Arial',12))
        main_table_text_pos_x = 70
        self.main_table_text.place(x=main_table_text_pos_x,y=50+dealt_y)

        #子表的listbox框
        self.sub_table_list=tk.Listbox(self.content_frame,height=14,width=text_width,font=('Arial',12))
        self.sub_table_list.place(x=main_table_text_pos_x,y=100+dealt_y)
        self.sub_table_list.bind('<Double-Button-1>',self.delete_sub_table)


        #提示
        lab=tk.Label(self.content_frame,text='*双击删除',fg='red')
        lab.place(x=main_table_text_pos_x, y=380+dealt_y)
        #选择子表的按钮
        self.select_sub_table_button = tk.Button(self.content_frame,height=1,text=' 选择子表 ',background='GhostWhite',command=self.select_sub_table)
        self.select_sub_table_button.place(x=600, y=130+dealt_y)
        #选择文件夹的按钮
        self.select_folder_button = tk.Button(self.content_frame,height=1,text='选择文件夹',background='GhostWhite')
        self.select_folder_button.place(x=600, y=230+dealt_y)
        #开始合并的按钮
        self.begin_meger_button = tk.Button(self.content_frame,height=1,text='开始合并',background='GhostWhite',command=self.begin)
        self.begin_meger_button.place(x=300,y=400+dealt_y)

        self.main_table_file = ''
        self.sub_table_file_list = ''
        #window.protocol('WM_DELETE_WINDOW',self.close_win)
    def delete_sub_table(self,event):
        pos = self.sub_table_list.curselection()
        self.sub_table_list.delete(pos)
        return

