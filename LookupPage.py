# encoding=utf-8
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
from lookup import *
import threading
from Enable_write import *
class LookupPage():
    def __init__(self,master):
        self.lookupPage=tk.Frame(master,width=700,height=500)
        self.lookupPage.pack()
        self.open_file_button1=tk.Button(self.lookupPage,text='打开表格',command=lambda :self.open_file1(self.file1),height=1)
        self.open_file_button1.place(x=50,y=50)
        self.file1=tk.Text(self.lookupPage,width=50,height=1,font=('Arial',12))
        self.file1.place(x=120,y=50)
        self.file1.delete('0.0','end')

        #是否跨表操作选项
        self.v1 = tk.BooleanVar()
        self.over_table=tk.Checkbutton(self.lookupPage,text='跨表操作',variable=self.v1,command=self.over_table_op)
        self.over_table.place(x=600,y=50)

        #如果跨表操作第二个表
        self.open_file_button2 = tk.Button(self.lookupPage, text='打开表格', command=lambda :self.open_file1(self.file2), height=1)
        self.open_file_button2.place(x=50, y=100)
        self.file2 = tk.Text(self.lookupPage,width=50, height=1, font=('Arial', 12))
        self.file2.place(x=120, y=100)
        self.file2.delete('0.0','end')
        #初始设置为不可见
        self.open_file_button2.place_forget()
        self.file2.place_forget()

        #源区域 返回区域 查找值区域 填充区域
        self.origin_lb=tk.Label(self.lookupPage,text='源区域:')
        self.origin_lb.place(x=50,y=180)
        self.origin_txt=tk.Text(self.lookupPage,width=10,height=1)
        self.origin_txt.place(x=120,y=180)

        self.return_lb = tk.Label(self.lookupPage, text='返回区域:')
        self.return_lb.place(x=200, y=180)
        self.return_txt = tk.Text(self.lookupPage, width=10, height=1)
        self.return_txt.place(x=270, y=180)

        self.lookup_lb = tk.Label(self.lookupPage, text='查找区域:')
        self.lookup_lb.place(x=50, y=220)
        self.lookup_txt = tk.Text(self.lookupPage, width=10, height=1)
        self.lookup_txt.place(x=120, y=220)

        self.write_lb = tk.Label(self.lookupPage, text='填充区域:')
        self.write_lb.place(x=200, y=220)
        self.write_txt = tk.Text(self.lookupPage, width=10, height=1)
        self.write_txt.place(x=270, y=220)

        #开始工作
        self.start_button=tk.Button(self.lookupPage,text='开始',command=self.start,width=8)
        self.start_button.place(x=500,y=190)

        #提示信息
        self.tips=tk.Label(self.lookupPage)
        self.tips.place(y=250,x=50)
        self.tips['text']='就绪'
        return

    def start(self):
        #使用分表合并
        if self.v1.get()==True:
            file1=self.file1.get('0.0','end').replace('\n','')
            file2=self.file2.get('0.0','end').replace('\n','')
            a=self.origin_txt.get('0.0','end').replace('\n','')
            b=self.return_txt.get('0.0','end').replace('\n','')
            c=self.lookup_txt.get('0.0','end').replace('\n','')
            d=self.write_txt.get('0.0','end').replace('\n','')
            if len(file1)==0 or len(file2)==0 or len(a)==0 or len(b)==0 or len(c)==0 or len(d)==0:
                tk.messagebox.showwarning(title='error', message='条件不完整')
                return
            thread1 = threading.Thread(target=begin,args=(file1,file2,a,b,c,d,self.tips))
            thread1.daemon = True
            thread1.start()
        else:
            file1 = self.file1.get('0.0', 'end').replace('\n','')
            a = self.origin_txt.get('0.0', 'end').replace('\n','')
            b = self.return_txt.get('0.0', 'end').replace('\n','')
            c = self.lookup_txt.get('0.0', 'end').replace('\n','')
            d = self.write_txt.get('0.0', 'end').replace('\n','')
            if len(file1)==0  or len(a)==0 or len(b)==0 or len(c)==0 or len(d)==0:
                tk.messagebox.showwarning(title='error', message='条件不完整')
                return
            thread1 = threading.Thread(target=begin, args=(file1,None, a, b, c, d,self.tips))
            thread1.daemon = True
            thread1.start()

        return

    def over_table_op(self):
        if self.v1.get()==True :
            self.open_file_button2.place(x=50, y=100)
            self.file2.place(x=120, y=100)
        elif self.v1.get()==False:
            self.open_file_button2.place_forget()
            self.file2.place_forget()
        return
    def open_file1(self,f):
        file_name=filedialog.askopenfilename(filetypes=[('excel file(.*xlsx)', '.xlsx')])
        file_name=file_name.replace('/','\\')
        if is_only_read(file_name):
            tkinter.messagebox.showwarning(title='error', message='主表文件已被某一应用程序打开，请关闭后重试')
            return
        f.delete('0.0','end')
        f.insert('insert',file_name)
        return

if __name__=='__main__':
    windows=tk.Tk()
    windows.geometry('700x500')
    LookupPage(windows)
    windows.mainloop()