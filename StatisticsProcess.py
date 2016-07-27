# -*- coding: gb2312 -*-
# import arcpy
# from arcpy.sa import *
# # todo.统计相关参数
#
# """
#     "越城":
# """
# sta = {}
# workspace =ur'D:\bulletinTemp\2015年\绍兴市'
# arcpy.env.workspace  = workspace
# infeature = u"D:/Program Files (x86)/LightningBulletin/LightningBulletin.gdb/绍兴市"
# ZonalStatisticsAsTable(infeature,'NAME',"lightningDay.tif",'stat_day',"","ALL")
# with arcpy.da.SearchCursor('stat_day',["NAME","MEAN","MIN","MAX"]) as cursor:
#     for row in cursor:
#         print row[0][:2],":",row[1],row[2],row[3]
#
# ZonalStatisticsAsTable(infeature,'NAME',"lightningDensity.tif",'stat_density',"","ALL")
# with arcpy.da.SearchCursor('stat_density',["NAME","MEAN","MIN","MAX"]) as cursor:
#     for row in cursor:
#          print row[0][:2],":",row[1],row[2],row[3]

# import arcpy
#
# arcpy.CreatePersonalGDB_management("/", "GDB.mdb")
# data = u'D:/bulletinTemp/2015年/data2015年.shp'
# arcpy.FeatureClassToGeodatabase_conversion(data,'/GDB.mdb')

import pyodbc
import win32com.client
import os

year = 2015
cwd = os.getcwd()  # 获取当前工作目录，便于程序移植
# 链接数据库
db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};',
                             'DBQ=', cwd, '/Data/GDB.mdb;']))  # Uid=Admin;Pwd=;')
cursor = db.cursor()

data_table = ''.join(['data',str(year),'年'])#sql查询语句不用使用中文

# 打开Excel应用程序
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
# 打开文件，即Excel工作薄
workbook = excel.Workbooks.Open(''.join([cwd, u'/Data/绍兴市雷电公报图表.xlsx']))

# ************浙江分地区统计**********
sql = """
SELECT count(*) AS num, Region
FROM %s
WHERE Province='浙江省'
GROUP BY Region
ORDER BY count(*) DESC
""" % data_table

results = {}
for row in cursor.execute(sql):
    results[row[1]] = row[0]  # 以 Region：num建立字典，方便下面赋值

sheet = workbook.Worksheets(u'省地区')
for row in xrange(2, 13):
    sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

# ********* 绍兴分县统计***********
sql = """
SELECT count(*) AS num, County
FROM %s
WHERE Region='绍兴市'
GROUP BY County
ORDER BY count(*) DESC
""" % data_table

results = {}
for row in cursor.execute(sql):
    results[row[1]] = row[0]  # 以 Region：num建立字典，方便下面赋值

sheet = workbook.Worksheets(u'市分区')
for row in xrange(3, 9):
    sheet.Cells(row, 3).Value = results[sheet.Cells(row, 2).Value]

# ************分月统计 月地闪次数和月平均强度(负闪)**************
sql = """
SELECT count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,1 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/1/1# And Date_< #YEAR/2/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,2 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/2/1# And Date_< #YEAR/3/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,3 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/3/1# And Date_< #YEAR/4/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,4 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/4/1# And Date_< #YEAR/5/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,5 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/5/1# And Date_< #YEAR/6/1#AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,6 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/6/1# And Date_< #YEAR/7/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,7 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/7/1# And Date_< #YEAR/8/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,8 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/8/1# And Date_< #YEAR/9/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,9 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/9/1# And Date_< #YEAR/10/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,10 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/10/1# And Date_< #YEAR/11/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,11 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/11/1# And Date_< #YEAR/12/1# AND Region='绍兴市' and Intensity<0)
union SELECT  count(*) as 负闪次数, -sum(Intensity)/count(*) as 平均强度 ,12 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/12/1# And Date_<=#YEAR/12/31# AND Region='绍兴市' and Intensity<0)
ORDER BY 月份
""".replace('QUERY_TABLE',data_table).replace('YEAR',str(year))

sheet1 = workbook.Worksheets(u'月份')
sheet2 = workbook.Worksheets(u'地闪强度波动')
i =1#行号
for row in cursor.execute(sql):
    i+=1
    sheet1.Cells(i,3).Value = row[0]#负闪次数
    sheet2.Cells(i,2).Value = row[1] if row[1] is not None else 0#负闪强度


# ************分月统计 月地闪次数和月平均强度(正闪)**************
sql = """
SELECT count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,1 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/1/1# And Date_< #YEAR/2/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,2 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/2/1# And Date_< #YEAR/3/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,3 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/3/1# And Date_< #YEAR/4/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,4 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/4/1# And Date_< #YEAR/5/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,5 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/5/1# And Date_< #YEAR/6/1#AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,6 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/6/1# And Date_< #YEAR/7/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,7 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/7/1# And Date_< #YEAR/8/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,8 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/8/1# And Date_< #YEAR/9/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,9 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/9/1# And Date_< #YEAR/10/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,10 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/10/1# And Date_< #YEAR/11/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,11 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/11/1# And Date_< #YEAR/12/1# AND Region='绍兴市' and Intensity>0)
union SELECT  count(*) as 正闪次数, sum(Intensity)/count(*) as 平均强度 ,12 as 月份
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/12/1# And Date_<=#YEAR/12/31# AND Region='绍兴市' and Intensity>0)
ORDER BY 月份
""".replace('QUERY_TABLE',data_table).replace('YEAR',str(year))

i =1#行号
for row in cursor.execute(sql):
    i+=1
    sheet1.Cells(i,2).Value = row[0] #正闪次数
    sheet2.Cells(i,3).Value = row[1] if row[1] is not None else 0 #正闪强度

workbook.Save()  # 保存EXCEL工作薄
workbook.Close()  # 关闭工作薄文件
excel.Quit()  # 关闭EXCEL应用程序
db.close()  # 关闭数据连接