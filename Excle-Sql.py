import pymssql
import xlrd
import xlwt
import time

start = time.clock()  # 开始计时
'''数据库连接'''
serverName = '127.0.0.1'
userName = 'sa'
password = 'hu12580WEI'
conn = pymssql.connect(serverName, userName, password, 'Gdky_Charge_BD_Formal')

'''uptdate函数用于合成sql语句和执行查询过程，检测是否执成功'''


def update(time, ID, sql_1, sql_2):  # time：修改的时间，ID：查询要修改时间的ID，sql_1：更新字段。sql_2：查询更新后的值作为检验
    cur = conn.cursor()
    sqls_1 = sql_1 % (time, ID)  # 连接sql语句，将每一次查询的ID和更改的时间整合到sql语句中
    cur.execute(sqls_1)  # 执行sql语句，更新ID对应的Cretetine字段
    conn.commit()  # 结束保存本次sql语句的执行
    cur.close()  # 关闭游标

    cur2 = conn.cursor()  # 新建一个游标
    sqls_2 = sql_2 % ID  # 连接sql语句，用于更新字段后查询这个ID对应更改的地方，返回这个字段
    cur2.execute(sqls_2)  # 执行sql语句
    result = cur2.fetchone()  # 查询后的结果赋值给result，返回的值为列表，格式为datetime
    conn.commit()  # 结束保存本次sql语句的执行
    '''检测是否更改要更改的内容，检测是否有这个ID，返回异常数据，输出到错误日志的表格'''
    if str(result) == 'None':  # 如果查询结果为空，代表数据库中没有该ID
        worksheet2.write(j, 0, ID)  # 将该ID写到表格中
    elif (str(result[0])[:19] != time[:19]):  # 否则如果不为空，检测查询到数据库中的更改字段的内容和表格中要更改的对比，查看是否更新
        worksheet2.write(j, 1, time)  # 如果不相同，将时间这个字段输出到错误日志表格中


workbook = xlrd.open_workbook("card.xlsx")  # 打开card.xlsx文件
sheet1 = workbook.sheet_by_index(0)  # 获取表格的第一页
nrows = sheet1.nrows  # 获取有效行数

workbook2 = xlwt.Workbook()  # 创建一个Excle工作表
worksheet2 = workbook2.add_sheet('sheet1')  # 创建一个sheet页，名字叫'sheet1'
worksheet2.write(0, 0, '未查询到的ID')  # 打印标头
worksheet2.write(0, 1, '错误时间')  # 打印标头
j = 1  # 标头占了第一行，从第二行开始记录错误日志
for i in range(nrows):  # 遍历表格每一行
    time_m = sheet1.cell_value(i, 1)  # 1代表的是第2列，第i行的数据,时间
    b = sheet1.cell_value(i, 0)  # 0代表的是第1列，第i行的数据,ID
    ID_m = int(b)  # 将浮点型b转换成整形（1.0装换为1，用于数据库查询ID）
    sql1 = "update  [Gdky_Charge_BD_Formal].[dbo].[t_Dat_CusCharges] set CreateTime ='%s' where ID='%d'"  # 更新字段的语句
    sql2 = "select CreateTime from [Gdky_Charge_BD_Formal].[dbo].[t_Dat_CusCharges]  where ID='%d'"  # 查询更改后的字段

    update(time_m, ID_m, sql1, sql2)  # 执行函数
    j = j + 1  # 写入错误日志中的数据，行数自增，逐行输出

conn.close()  # 关闭连接
workbook2.save('Excel_ErrorLog.xls')  # 保存错误表格
print('Tips:成功更改')
end = time.clock()
print('本次用时: %s 秒' % (end - start))
print('完成')
