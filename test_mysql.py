# encoding=utf-8
import MySQLdb

host = 'localhost'
port = 3306
user = 'root'
passwd = 'root'
db = 'temp_parse_word'
conn = MySQLdb.connect(host, user, passwd, db)
cursor = conn.cursor()

cursor.execute("select * from device_base_info")
data = cursor.fetchone()

print(data)
#cur.close() 关闭游标
cursor.close()

#conn.commit()方法在提交事物，在向数据库插入一条数据时必须要有这个方法，否则数据不会被真正的插入。
conn.commit()

#conn.close()关闭数据库连接
conn.close()