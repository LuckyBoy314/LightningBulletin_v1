# -*- coding: gb2312 -*-
# import arcpy
# from arcpy.sa import *
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
# arcpy.CreatePersonalGDB_management(''.join([cwd,'/data']), "GDB.mdb")
# data = u'D:/bulletinTemp/2015年/data2015年.shp'
# arcpy.FeatureClassToGeodatabase_conversion(data,''.join([cwd,'/data','/GDB.mdb']))

import pyodbc
import win32com.client
import os

year = 2015
cwd = os.getcwd()  # 获取当前工作目录，便于程序移植
# 链接数据库
db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};',
                             'DBQ=', cwd, '/Data/GDB.mdb;']))  # Uid=Admin;Pwd=;')
cursor = db.cursor()

data_table = ''.join(['data', str(year), '年'])  # sql查询语句不用使用Unicode

# 打开Excel应用程序
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
# 打开文件，即Excel工作薄
workbook = excel.Workbooks.Open(''.join([cwd, u'/Data/公报图表模板.xlsx']))

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

sheet = workbook.Worksheets(u'省分区统计')
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

sheet = workbook.Worksheets(u'市分区统计')
for row in xrange(2, 8):
    sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

# todo SQL查询有待优化
# ************分月统计 月地闪次数和月平均强度(负闪)**************
sql = """
SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,1 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,2 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,3 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,4 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,5 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,6 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,7 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,8 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,9 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,10 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,11 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
union SELECT  count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,12 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
ORDER BY 月份
""".replace('QUERY_TABLE', data_table).replace('YEAR', str(year))

sheet = workbook.Worksheets(u'分月统计')
i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 2).Value = row[0]  # 负闪次数
    sheet.Cells(i, 5).Value = row[1] if row[1] is not None else 0  # 负闪强度

# ************分月统计 月地闪次数和月平均强度(正闪)**************
sql = """
SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,1 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,2 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,3 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,4 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,5 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#AND Region='绍兴市'
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,6 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,7 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,8 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,9 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,10 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,11 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,12 AS 月份
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
ORDER BY 月份
""".replace('QUERY_TABLE', data_table).replace('YEAR', str(year))

i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # 正闪次数
    sheet.Cells(i, 6).Value = row[1] if row[1] is not None else 0  # 正闪强度

# ************分时段统计 时段地闪次数和时段平均强度(负闪)**************
sql = """
SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,0 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=0
UNION SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,1 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=1
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,2 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=2
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,3 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=3
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,4 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=4
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,5 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=5
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,6 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=6
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,7 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=7
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,8 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=8
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,9 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=9
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,10 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=10
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,11 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=11
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,12 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=12
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,13 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=13
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,14 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=14
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,15 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=15
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,16 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=16
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,17 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=17
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,18 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=18
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,19 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=19
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,20 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=20
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,21 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=21
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,22 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=22
union SELECT count(*) AS 负闪次数, -sum(Intensity)/count(*) AS 平均强度 ,23 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Val(Time_)=23
ORDER BY 时段
""".replace('QUERY_TABLE', data_table)

sheet = workbook.Worksheets(u'分时段统计')
i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 2).Value = row[0]  # 负闪次数
    sheet.Cells(i, 5).Value = row[1] if row[1] is not None else 0  # 负闪强度

# ************分时段统计 时段地闪次数和时段平均强度(正闪)**************
sql = """
SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,0 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=0
UNION SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,1 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=1
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,2 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=2
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,3 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=3
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,4 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=4
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,5 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=5
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,6 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=6
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,7 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=7
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,8 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=8
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,9 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=9
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,10 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=10
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,11 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=11
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,12 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=12
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,13 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=13
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,14 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=14
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,15 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=15
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,16 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=16
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,17 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=17
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,18 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=18
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,19 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=19
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,20 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=20
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,21 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=21
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,22 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=22
union SELECT count(*) AS 正闪次数, sum(Intensity)/count(*) AS 平均强度 ,23 AS 时段
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Val(Time_)=23
ORDER BY 时段
""".replace('QUERY_TABLE', data_table)

i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # 正闪次数
    sheet.Cells(i, 6).Value = row[1] if row[1] is not None else 0  # 正闪强度

# **********负闪强度分布**************
sql = """
SELECT count(*) AS 负闪次数,0 AS 左边界,5 AS 右边界
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=0 AND Abs(Intensity)<5
union SELECT count(*) AS 负闪次数,5,10
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=5 AND Abs(Intensity)<10
union SELECT count(*) AS 负闪次数,10,15
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=10 AND Abs(Intensity)<15
union SELECT count(*) AS 负闪次数,15,20
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=15 AND Abs(Intensity)<20
union SELECT count(*) AS 负闪次数,20,25
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=20 AND Abs(Intensity)<25
union SELECT count(*) AS 负闪次数,25,30
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=25 AND Abs(Intensity)<30
union SELECT count(*) AS 负闪次数,30,35
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=30 AND Abs(Intensity)<35
union SELECT count(*) AS 负闪次数,35,40
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=35 AND Abs(Intensity)<40
union SELECT count(*) AS 负闪次数,40,45
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=40 AND Abs(Intensity)<45
union SELECT count(*) AS 负闪次数,45,50
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=45 AND Abs(Intensity)<50
union SELECT count(*) AS 负闪次数,50,55
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=50 AND Abs(Intensity)<55
union SELECT count(*) AS 负闪次数,55,60
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=55 AND Abs(Intensity)<60
union SELECT count(*) AS 负闪次数,60,65
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=60 AND Abs(Intensity)<65
union SELECT count(*) AS 负闪次数,65,70
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=65 AND Abs(Intensity)<70
union SELECT count(*) AS 负闪次数,70,75
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=70 AND Abs(Intensity)<75
union SELECT count(*) AS 负闪次数,75,80
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=75 AND Abs(Intensity)<80
union SELECT count(*) AS 负闪次数,80,85
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=80 AND Abs(Intensity)<85
union SELECT count(*) AS 负闪次数,85,90
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=85 AND Abs(Intensity)<90
union SELECT count(*) AS 负闪次数,90,95
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=90 AND Abs(Intensity)<95
union SELECT count(*) AS 负闪次数,95,100
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=95 AND Abs(Intensity)<100
union SELECT count(*) AS 负闪次数,100,150
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=100 AND Abs(Intensity)<150
union SELECT count(*) AS 负闪次数,150,200
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=150 AND Abs(Intensity)<200
union SELECT count(*) AS 负闪次数,200,250
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=200 AND Abs(Intensity)<250
union SELECT count(*) AS 负闪次数,250,300
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=250 AND Abs(Intensity)<300
UNION SELECT count(*) AS 负闪次数,300,1000
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity<0 AND Abs(Intensity)>=300
ORDER BY 左边界
""".replace("QUERY_TABLE", data_table)

sheet = workbook.Worksheets(u'强度分布统计')
i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # 负闪次数

# ***********正闪强度分布************
sql = """
SELECT count(*) AS 正闪次数,0 AS 左边界,5 AS 右边界
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=0 AND Intensity<5
union SELECT count(*) AS 正闪次数,5,10
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=5 AND Intensity<10
union SELECT count(*) AS 正闪次数,10,15
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=10 AND Intensity<15
union SELECT count(*) AS 正闪次数,15,20
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=15 AND Intensity<20
union SELECT count(*) AS 正闪次数,20,25
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=20 AND Intensity<25
union SELECT count(*) AS 正闪次数,25,30
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=25 AND Intensity<30
union SELECT count(*) AS 正闪次数,30,35
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=30 AND Intensity<35
union SELECT count(*) AS 正闪次数,35,40
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=35 AND Intensity<40
union SELECT count(*) AS 正闪次数,40,45
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=40 AND Intensity<45
union SELECT count(*) AS 正闪次数,45,50
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=45 AND Intensity<50
union SELECT count(*) AS 正闪次数,50,55
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=50 AND Intensity<55
union SELECT count(*) AS 正闪次数,55,60
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=55 AND Intensity<60
union SELECT count(*) AS 正闪次数,60,65
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=60 AND Intensity<65
union SELECT count(*) AS 正闪次数,65,70
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=65 AND Intensity<70
union SELECT count(*) AS 正闪次数,70,75
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=70 AND Intensity<75
union SELECT count(*) AS 正闪次数,75,80
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=75 AND Intensity<80
union SELECT count(*) AS 正闪次数,80,85
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=80 AND Intensity<85
union SELECT count(*) AS 正闪次数,85,90
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=85 AND Intensity<90
union SELECT count(*) AS 正闪次数,90,95
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=90 AND Intensity<95
union SELECT count(*) AS 正闪次数,95,100
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=95 AND Intensity<100
union SELECT count(*) AS 正闪次数,100,150
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=100 AND Intensity<150
union SELECT count(*) AS 正闪次数,150,200
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=150 AND Intensity<200
union SELECT count(*) AS 正闪次数,200,250
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=200 AND Intensity<250
union SELECT count(*) AS 正闪次数,250,300
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=250 AND Intensity<300
UNION SELECT count(*) AS 正闪次数,300,1000
FROM QUERY_TABLE
WHERE Region='绍兴市' AND Intensity>=300
ORDER BY 左边界
""".replace("QUERY_TABLE", data_table)

i = 1  # 行号
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 4).Value = row[0]  # 正闪次数

workbook.Save()  # 保存EXCEL工作薄
workbook.Close()  # 关闭工作薄文件
excel.Quit()  # 关闭EXCEL应用程序
db.close()  # 关闭数据连接