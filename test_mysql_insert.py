# encoding=utf-8
import MySQLdb

host = 'localhost'
port = 3306
user = 'root'
passwd = 'root'
db = 'temp_parse_word'
conn = MySQLdb.connect(host, user, passwd, db)
cursor = conn.cursor()

data = {
    'id': '2',
}
table = 'device_base_info'
keys = ', '.join(data.keys())
values = ', '.join(['%s'] * len(data))
sql = 'INSERT INTO {table}({keys}) VALUES ({values})'.format(table=table, keys=keys, values=values)
try:
   cursor.execute(sql, tuple(data.values()))
   print('Successful')
   conn.commit()
except:
   print('Failed')
   conn.rollback()
cursor.close()
conn.close()