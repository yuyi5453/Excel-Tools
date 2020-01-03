#coding=gbk
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


try:
    try:
        raise SystemError
    except:
        print('内层ex')
    finally:
        print('内层')
except:
    print('外层ex')
finally:
    print('外层')
