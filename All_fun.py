# coding=gbk
# 此文件包括公式生成功能的各种函数

# 将【C】与【A】进行匹配，并将对应【B】列中的值作为结果填入【D】中

def lookup(wb,A,B,C,D):
    sht=wb.sheets[0]
    A_rng=sht.range(A).expand('down').value
    B_rng=sht.range(B).expand('down').value
    C_rng=sht.range(C).expand('down').value
    D_rng=list()

    for c in C_rng:
        flag=0
        for i in range(0,len(A_rng)):
            a = A_rng[i]
            if c==a:
                flag=1
                D_rng.append(B_rng[i])
                break

        if flag==0:
            D_rng.append('None')

    sht.range(D).options(transpose=True).value=D_rng
    return

import xlwings as xl
try:
    app = xl.App(visible=False)
    wb = app.books.open(r'C:\Users\DSJ\Desktop\test.xlsx')
    A = 'A2:A8'
    B = 'B2:B8'
    C = 'C2:C4'
    D = 'D2:D4'
    lookup(wb, A, B, C, D)
    wb.save()
    wb.close()
finally:
    app.quit()