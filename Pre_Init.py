#coding=gbk
import xlwings as xw
def Pre_Init(wb):
    sht=wb.sheets[0];
    rng=sht.range("A1").expand("down").value
    l=len(rng)
    if (isinstance(rng, str) is True):
        l=1
    print(rng,l)
    right_rng=sht.range((l,1)).expand("right").value
    print(right_rng)
    col_num=len(right_rng)#列个的数

    return col_num,l+1
