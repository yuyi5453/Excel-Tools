#文件在被其他应用打开时，是只读的，无法被写入
#函数功能为判断文件是否只读，提示使用者关闭文件
def is_only_read(file_name):
    try:
        with open(file_name, "r+") as fr:
            return False
    except IOError as e:
        if "[Errno 13] Permission denied" in str(e):
            return True
        else:
            print(str(e))
            return False