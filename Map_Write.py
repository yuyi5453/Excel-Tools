#将mem_map中的数据存入到wb中
def map_write(mem_map,wb,row_start,col_num):

    all_key=mem_map.keys()
    sht=wb.sheets[0]
    begin=row_start
    list = []
    for key in all_key:
        li = mem_map[key]
        list.append(li)
    sht.range((begin,1)).value = list
    return