# encoding:utf-8
import copy

def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False

def std_handle(li, key):
    value = li[key - 1]
    if (isinstance(value, str)):
        li[key-1]=value.replace(' ','')
    elif (isinstance(value, float)):
        if (int(value) == value):
            li[key - 1] = str(int(value))
        else:
            li[key - 1] = str(value)

    return


def map_add(mem_map, li, key):
    mem_map[li[key - 1]] = []
    mem_map[li[key - 1]] = copy.deepcopy(li)
    print(mem_map)
    return


def map_modify(mem_map, lis, key,need_add_col):
    # mem_map是一个字典，list是列表，key表示哪一列是作为键值列
    key_value = lis[key - 1]
    length = len(lis)

    #把有可能为字符串的可加列转换为数字
    for c in need_add_col:
        if lis[c-1] is None:
            continue
        if is_number(lis[c-1]) == False:
            print('可加列输入错误')
            raise SystemExit
        if isinstance(lis[c-1],str):
            lis[c-1]=float(lis[c-1].replace(' ',''))

    for i in range(length):
        if (i == key - 1):
            continue
        if (mem_map[key_value][i] == None):
            mem_map[key_value][i] = copy.deepcopy(lis[i])

        # elif i+1 in tuple
        elif (i+1 in need_add_col):
            mem_map[key_value][i] = lis[i] + mem_map[key_value][i]

    print(mem_map)
    return