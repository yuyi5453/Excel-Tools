# encoding:utf-8
import tkinter as tk
from meger_ui import *
from LookupPage import *

def merger_button_onClick(meger_content):
    lookup_content.lookupPage.place_forget()
    meger_content.content_frame.place(x=100,y=0)

def lookup_button_onClick(lookup_content):
    meger_content.close_meger_content()
    lookup_content.lookupPage.place(x=100,y=0)

if __name__ == '__main__':
    window = tk.Tk()
    window.geometry('800x500')
    #navigation_frame: 左边的导航界面
    navigation_frame = tk.Frame(window,width=100,height=500,bg='green')
    navigation_frame.place(x=0,y=0)
    #获得合并界面的实例
    meger_content = merger_UI(window)
    lookup_content=LookupPage(window)
    #布置导航界面的图片
    canvas = tk.Canvas(navigation_frame, bg='blue', height=500, width=100)
    canvas.place(x=0,y=0)
    image_file = tk.PhotoImage(file='img\p1.jpg')
    image = canvas.create_image(0, 0, anchor='nw', image=image_file)
    #导航界面的lookup按钮和合并按钮
    lookup_button = tk.Button(navigation_frame,height=2,text='lookup\n函数生成',background='GhostWhite',command=lambda:lookup_button_onClick(lookup_content))
    lookup_button.place(x=20,y=130)
    meger_button = tk.Button(navigation_frame,height=2,text='合并表格',
                              background='GhostWhite',command = lambda arg=meger_content:merger_button_onClick(arg))
    meger_button.place(x=20,y=260)
    window.mainloop()