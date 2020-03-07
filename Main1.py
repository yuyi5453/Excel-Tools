# coding=gbk
import sys
import xlwings as xw
from Pre_Init import *
from All_Merge import *

# main_table_file=input()
# key=input()
# sub_table_file_list=[]
# for line in sys.stdin:
#     sub_table_file_list.extend(line.split())
# main_table_file=r'G:\计算机设计大赛\测试文件\模板(母表).xlsx'
# # key=2
# # sub_table_file_list=[
# # r'G:\计算机设计大赛\测试文件\子表1_B.xlsx',
# # r'G:\计算机设计大赛\测试文件\子表2_软二CD.xlsx',
# # r'G:\计算机设计大赛\测试文件\子表3_软一CD.xlsx',
# # r'G:\计算机设计大赛\测试文件\子表4_软一2月A.xlsx',
# # r'G:\计算机设计大赛\测试文件\子表5_软一3月A .xlsx',
# # r'G:\计算机设计大赛\测试文件\子表6_软二2月A .xlsx',
# # r'G:\计算机设计大赛\测试文件\子表7_软二3月A.xlsx']
# # need_add=(4,)


# need_add=tuple(input())
def begin(main_table_file,sub_table_file_list):
    # main_table_file=r'G:\计算机设计大赛\测试文件2\m1.xlsx'
    key=2
    need_add=()
    print(main_table_file)
    print(sub_table_file_list)
    for name in sub_table_file_list:
        print(name)
    # sub_table_file_list=[
    # r'G:\计算机设计大赛\测试文件2\z1.xlsx',
    # r'G:\计算机设计大赛\测试文件2\z2.xlsx',
    # r'G:\计算机设计大赛\测试文件2\z3.xlsx',
    # r'G:\计算机设计大赛\测试文件2\z4.xlsx',
    # r'G:\计算机设计大赛\测试文件2\z5.xlsx'
    # ]
    # 试试

    app=xw.App(visible=False)
    main_wb=app.books.open(main_table_file)

    mem_map={}
    try:
        print('预处理')
        col_num,row_start=Pre_Init(wb=main_wb)
        print('开始合并')
        all_merge(wb=main_wb,row_start=row_start,col_num=col_num,file_name_list=sub_table_file_list,mem_map=mem_map,key=key,need_add_col=need_add)
        print('合并完成')

    finally:
        print('finally')
        main_wb.save()
        main_wb.close()
        app.quit()
        app.kill()
