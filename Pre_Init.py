#coding=gbk
import xlwings as xw
def Pre_Init(wb):
    sht=wb.sheets[0];
    rng=sht.range("A1").expand("down").value
    if rng==None: #rng=None 代表母表没有格式
        col_num=0
        l=0
    else: #母表有格式
        l=len(rng)
        if  isinstance(rng, str) is True:
            l=1
        right_rng=sht.range((l,1)).expand("right").value
        col_num = len(right_rng)  # 列个的数
        if isinstance(right_rng,str):
            col_num=1

    return col_num,l+1
