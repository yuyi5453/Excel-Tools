# encoding:utf-8
import xlwings as xw
import numpy as np
import wx
import copy
app=xw.App(visible=False,add_book=False)
app.display_alerts=False

wb=app.books.open(r'G:\计算机设计大赛\111.xlsx')
sht=wb.sheets[0]
rng=sht.range("A1")
rng.value='123   '


print(type(rng),rng)

wb.save()
wb.close()
app.quit()
app.kill()
