#coding=gbk
# Excel-Tools


存在的bug：
母表没有格式的时候，预处理的rng回返回None,然后抛出异常.(解决)
母表没有格式的时候，无法得出表的宽度，即col_num