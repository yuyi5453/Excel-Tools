
# encoding:utf-8
import sys
import xlwings as xw
from Pre_Init import *
from All_Merge import *

def begin(main_table_file,sub_table_file_list,key,need_add,label2):
    try:
        app=xw.App(visible=False)
        main_wb=app.books.open(main_table_file)

        mem_map={}

        label2.config(text='预处理')
        col_num,row_start=Pre_Init(wb=main_wb)
        label2.config(text='开始合并')
        all_merge(wb=main_wb,row_start=row_start,col_num=col_num,file_name_list=sub_table_file_list,mem_map=mem_map,key=key,need_add_col=need_add,label2=label2)
        label2.config(text='合并完成')
        main_wb.save()
        main_wb.close()

    finally:
        app.quit()
        app.kill()


