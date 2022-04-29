#coding=utf8
import xlrd
import MySQLdb

def open_excel(file = 'file.xls') :
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))

    # 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引


def excel_table_byindex(data='', by_index=0):
    table = data.sheets()[by_index]
    sheet_names = data.sheet_names()
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    list = []
    for rownum in range(2, nrows):
        mysqlData = {
            'sheet_name': sheet_names[by_index],
            'sort_no': table[rownum][0].value,
            'device_name': table[rownum][1].value,
            'device_code': table[rownum][2].value,
            'host_name': table[rownum][3].value,
            'device_important_score': table[rownum][4].value,
            'device_reliability_score': table[rownum][5].value,
            'device_running_time_score': table[rownum][6].value,
            'device_running_environment_score': table[rownum][7].value,
            'device_easy_maintain_score': table[rownum][8].value,
            'device_maintain_frequency': table[rownum][9].value,
        }
        list.append(mysqlData)
    return list


def main():
    excel_name = 'C:\\Users\\yanpe\\Desktop\\设备评分-AFC监视中心-20220424.xls'
    data = open_excel(excel_name)
    lists = []
    for i in range (1, len(data.sheets())):
        list = excel_table_byindex(data, i)
        lists.append(list)
    mysql_data=deal_data(lists)
    print(mysql_data)
    mysql_insert(mysql_data)

def mysql_insert(mysql_list=None):
    host = 'localhost'
    port = 3306
    user = 'root'
    passwd = 'root'
    db = 'temp_parse_word'
    conn = MySQLdb.connect(host, user, passwd, db)
    cursor = conn.cursor()
    for data in range (0, len(mysql_list)):
        data_dict = dict(mysql_list[data])
        try:
            table = 'device_score_excel'
            keys = ', '.join(data_dict.keys())
            values = ', '.join(['%s'] * len(data_dict.keys()))
            sql = 'INSERT INTO {table}({keys}) VALUES ({values})'.format(table=table, keys=keys, values=values)
            cursor.execute(sql, tuple(data_dict.values()))
            print('Successful')
            conn.commit()
        except Exception as err:
            print('mysql error ', err)
            conn.rollback()

def deal_data(lists=None):
    if lists is None:
        lists = []
    mysql_list=[]
    for list in lists:
        for data in list:
            mysql_list.append(data)
    return mysql_list

if __name__ == "__main__":
    main()

