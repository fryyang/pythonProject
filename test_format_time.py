import re

def time_format(str):
    if '/' in str:
        if str.count('/') == 1:
            str = str + '/01'
        return str
    if '年' in str:
        str = re.sub(r'年', "/", str)
    if '月' in str:
        str = re.sub(r'月', "/", str)
        if '日' in str:
            str = re.sub(r'日', "", str)
        else:
            str = str + '01'
    return str

print(time_format('2022/10'))
