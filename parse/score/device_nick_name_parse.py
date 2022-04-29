# coding=utf-8
from docx import Document
import MySQLdb

host = 'localhost'
port = 3306
user = 'root'
passwd = 'root'
db = 'temp_parse_word'
conn = MySQLdb.connect(host, user, passwd, db)

#读取文档
filename = "C:\\Users\\yanpe\\Desktop\\20211015.docx"
doc = Document(filename) #filename为word文档

cursor = conn.cursor()
#获取文档中的表格
i = 101
for p in doc.paragraphs:
	style_name = p.style.name
	if style_name.startswith('Heading 4'):
		# print(style_name, p.text, sep=':')
		sql = "update device_base_info set device_nick_name ='" + p.text + "' where id=" + str(i)
		try:
			print(sql)
			cursor.execute(sql)
			print('Successful')
		except Exception as err:
			print('mysql error ', err)
			conn.rollback()
		i = i + 1

#cur.close() 关闭游标
cursor.close()

#conn.commit()方法在提交事物，在向数据库插入一条数据时必须要有这个方法，否则数据不会被真正的插入。
conn.commit()

#conn.close()关闭数据库连接
conn.close()