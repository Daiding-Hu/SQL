import pymssql
import xlwt

'''连接本地数据库相关配置信息'''
serverName = '127.0.0.1'  # 数据库地址，本机用127.0.0.1
userName = 'sa'  # 登录名
passWord = 'hu12580WEI'  # 密码

'''连接数据库并获取cursor'''
conn = pymssql.connect(serverName, userName, passWord, 'Gdky_Charge_BD_Formal')

cursor = conn.cursor(as_dict=True)  # 如果指定了as_dict为True，则返回结果变为字典类型
args = 'U'  # 组成sql语句中的变量，用于替换sql语句中的表名
sql = "select name from sysobjects where xtype='%s'"  # 要执行的sql语句，%s为可以替换的内容
T = sql % args  # 组合后复制给T
cursor.execute(T)  # 执行查询语句
list_name = cursor.fetchall()  # 返回所有的查询结果给list_name
conn.commit()  # 结束本次查询
table_name = []  # 新建一个列表
for dic in list_name:  # 遍历查询结果list_name中的每一个字典元素
    for i in dic:  # 遍历上一个循环中选中的字典内的元素
        table_name.append(dic[i])  # 将字典中元素的键值返回，并逐个添加到table_name中


# 通过控制args变量（不同的表名称）输出不同的查询结果
def select(args, sql_t):
    cursor = conn.cursor()
    T = sql_t % args  # 合成sql语句
    cursor.execute(T)  # 执行查询
    list_name_t = cursor.fetchall()  # 内容赋给变量，返回类型为列表
    conn.commit()  # 结束本次查询
    return list_name_t  # 返回列表


'''写入表格操作'''
workbook = xlwt.Workbook()  # 创建一个Excle工作表
worksheet = workbook.add_sheet('sheet1')  # 创建一个sheet页，名字叫'sheet'
n = 1
m = 0
k = 1
j = 0
title = [u'表名', u'字段名', u'字段类型']  # 表格中的标题需要的内容放入列表
for i in title:  # 以此输入到表格中
    worksheet.write(0, j, i)
    j = j + 1
'''上面已经得到了一个table_name列表，遍历这个列表中的每一个表名，通过select函数查询每一个表名中含有什么字段'''
for table_t in table_name:  # 遍历得到表名
    sql_t = "select sc.name,st.name from syscolumns sc,systypes st where sc.xtype=st.xtype and sc.id in(select id from sysobjects where xtype='U' and name='%s')"
    Sql_Result = select(table_t, sql_t)  # 将结果给Sql_Result
    a = table_t  # 遍历得到的表名给变量a
    m = len(Sql_Result) + n - 1  # 列表的长度，用于控制合并多少个单元格
    worksheet.write_merge(n, m, 0, 0, a)  # 合并单元格操作，第n行到第m行合并，0,0表示的是列，a表示合并后单元格中的内容
    n = m + 1  # 循环控制下一个合并开始的位置
    for i in Sql_Result:
        worksheet.write(k, 1, i[0])  # 写入查询到的字段名称i[0]，因为第一行（下标为0的行 ）已经用做标题，k从1开始，按行输入到第2列
        worksheet.write(k, 2, i[1])  # 写入查询到的字段类型i[2]，因为第一行（下标为0的行 ）已经用做标题，k从1开始，按行输入到第三列
        k = k + 1

workbook.save('Excel_test.xls')
print('Tips：成功导出Excle')
