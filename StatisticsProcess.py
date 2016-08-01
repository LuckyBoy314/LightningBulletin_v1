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

#参数
year = 2015
target_area = u'绍兴市'
area = 8256.0
province_area = {u'杭州市':16596.0,u'宁波市':9365.0,u'温州市':11784.0,u'湖州市':5794.0,u'嘉兴市':3915.0,
u'绍兴市':8256.0,u'金华市':10919.0,u'台州市':9413.0,u'舟山市':1440.0,u'衢州市':8837.0,u'丽水市':17298.0}
#todo 各地区面积要包括海域面积

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

#处理SQL查询结果，顺便记录本地区地闪次数在全省的排名
results = {}
rank = 0
for row in cursor.execute(sql):
    rank+=1
    results[row[1]] = row[0]  # 以 Region：num建立字典，方便下面赋值
    if row [1] == target_area:
        sum_rank_in_province = rank #本地区地闪次数在全省的排名

#将SQL查询结果写入Excel
sheet = workbook.Worksheets(u'省分区统计')
for row in xrange(2, 13):
    sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

sum_region = results[target_area]  #本地区地闪总数
density_region = sum_region/area  #本地区地闪密度

#计算全省各地区地闪密度和全省平均密度
density_province_dict = {}
density_province = 0
for key in results:
    density_province_dict[key] = results[key]/province_area[key]
    density_province+= density_province_dict[key]
density_province/=len(province_area) #全省平均地闪密度
#密度从大到小进行排序
density_province_sorted = sorted(density_province_dict.iteritems(),key = lambda d:d[1],reverse=True)
#计算本地区地闪密度排名
rank = 0
for item in density_province_sorted:
    rank+=1
    if item[0] == target_area:
        density_rank_in_province = rank #本地区地闪密度在全省的排名
        break

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




if 0.05<=(sum_region-sum_region_lastyear)/float(sum_region)<0.1:
    compare_with_lastyera = u'略有增多'
elif -0.1<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.05:
    compare_with_lastyera = u'略有较少'
elif 0.1<=(sum_region-sum_region_lastyear)/float(sum_region)<0.5:
    compare_with_lastyera = u'有所增多'
elif -0.5<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.1:
    compare_with_lastyera = u'有所较少'
elif 0.5<=(sum_region-sum_region_lastyear)/float(sum_region)<0.9:
    compare_with_lastyera = u'增幅较大'
elif -0.9<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.5:
    compare_with_lastyera = u'减幅较大'
elif 0.9<=(sum_region-sum_region_lastyear)/float(sum_region):
    compare_with_lastyera = u'大幅增多'
elif (sum_region-sum_region_lastyear)/float(sum_region)<=-0.9:
    compare_with_lastyera = u'大幅减少'
else:
    compare_with_lastyera = u'基本持平'

if density_region>density_province:
	compare_with_province = u'高于'
else:
	compare_with_province = u'低于'

p1 = u'%d年我市共发生地闪%d次，平均地闪密度%.2f次/km?，平均雷暴日%d天（见表1-1）。 \
     与上年的地闪%d次相比，%s。从时间分布来看，地闪主要集中在%s[6、7、8]月，\
     三个月地闪占全年总地闪次数的%.2f%。从空间分布来看，%s发生地闪次数最多，%s最少。\
     全市地闪平均密度%s全省平均的%.2f次/km?，在全省各市中%s闪次数排第%s位，\
     地闪平均密度排第%s位（见表1-2）。'% (year,sum_region,density_region,day_region,
                            sum_region_lastyear, compare_with_lastyear, months_most,months_percent,
                            county_most,county_least,compare_with_province,density_province,
                            target_area,sum_rank_in_province,density_rank_in_province)

p2 = u'据不完全统计，2016年全市因雷电引发的灾害共148起，无人员伤亡事故。\
     造成直接经济损失达7788.04万元，间接经济损失677.42万元。'

p3= u'从地区统计来看，地区分布相对不均，诸暨地闪次数最多，共4788次，越城区最少，只有106次，\
    两者分别占全市总地闪数的51.64%和1.14%。从平均密度统计来看，诸暨密度最高，为2.07次/km2，\
    越城区最低，为0.21次/km2（见表1-1）。'

p4 = u'从地闪密度空间分布图上（见图1-1）可以看出，诸暨西北部、嵊州和诸暨交界区域地闪密度较高，\
     最高超过5次/km2。新昌东部有部分地区，地闪密度超过3次/km2，全市大部分地区地闪密度小于2次/km2。'

p5= u'现行国家标准所引用的雷暴日指人工观测（测站周围约15km半径域面）有雷暴天数的多年平均。\
    根据我省闪电定位监测资料推算（以15km为间隔，分别统计各点15km半径范围内的雷暴日，再插值推算），\
    2016年全市地闪雷暴日平均43天，最低为29天，最高67天。空间分布上来看，北部平原地区雷暴日较少，\
    西南大部和东南部分区域雷暴天数较多（见图1-2）。'

p6 = u'2016年绍兴市雷电初日为2月25日。从分月统计来看，地闪次数随月份呈现近似正态分布特征，1月和12月未监测到地闪，\
     地闪次数峰值出现在8月，6-8月是雷暴高发的月份 ，三个月地闪次数占总数的%d。正、负地闪平均强度的峰值分别在2月和3月，\
     其他月份波动平缓(见图1-3) 。'

p7 = u'从分时段统计来看，地闪次数峰值出现在第18个时段（17:00-18:00），地闪主要集中在午后两点到晚上九点，\
     七个时段内的地闪次数占总数的d%。地闪平均强度随时间呈波状起伏特征，但总体波动不大。\
     正闪平均强度峰值在第7个时段（7:00-8:00），负闪平均强度峰值在第11个时段（11:00-12:00）(见图1-4)。'

p8 = u'由正、负地闪强度分布图可见，地闪次数随地闪强度呈近似正态分布特征。正地闪主要集中在5-60kA内（见图1-5），\
     该区间内正地闪次数约占总地闪的87.20%，负地闪主要分布在5-60kA内（见图1-6），\
     该区间内负地闪次数约占总负地闪的91.46%。'

workbook.Save()  # 保存EXCEL工作薄
workbook.Close()  # 关闭工作薄文件
excel.Quit()  # 关闭EXCEL应用程序
db.close()  # 关闭数据连接
