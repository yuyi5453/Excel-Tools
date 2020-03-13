# encoding:utf-8
#合并所有表到map中
from Sub_Table_Solve import *
from  Map_Write import *
import xlwings as xw
def all_merge(mem_map,wb,file_name_list,row_start,col_num,key,need_add_col,label2):
    for file_name in file_name_list:
        # 打开子文件
        label2.config(text='打开文件:'+file_name)
        try:
            app=xw.App(visible=False)
            sub_wb=app.books.open(file_name)
            # 子文件处理
            sub_table_solve(mem_map=mem_map,sub_wb=sub_wb,row_start=row_start,col_num=col_num,key=key,need_add_col=need_add_col)
        finally:
            #关闭子文件
            label2.config(text=file_name+'处理完成')
            sub_wb.save()
            sub_wb.close()
            app.quit()
            app.kill()
    # 将map中的数据写到excel
    label2.config(text= '正在写入')
    map_write(mem_map=mem_map,wb=wb,row_start=row_start,col_num=col_num)
    return