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

#链接数据库
#db = pyodbc.connect("DSN=<that Data Source I just created>")
db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};  DBQ=D:/GDB.mdb;')# Uid=Admin;Pwd=;')
cursor = db.cursor()
data_table = "data2015年"

#打开Excel应用程序
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
#打开文件，即Excel工作薄
workbook = excel.Workbooks.Open(ur'D:\Program Files (x86)\LightningBulletin\Data\绍兴市雷电公报图表.xlsx')


#浙江分地区统计
sql_regions = """
SELECT count(*) AS num, Region
FROM %s
WHERE Province='浙江省'
GROUP BY Region
ORDER BY count(*) DESC;
"""%data_table

results = {}
for row in cursor.execute(sql_regions):
    results[row[1]] = row[0]# 以 Region：num建立字典，方便下面赋值

sheet_regions = workbook.Worksheets(u'省地区')
for row in xrange(2,13):
    region = sheet_regions.Cells(row,2).Value
    sheet_regions.Cells(row,1).Value = results[region]


#绍兴分县统计
sql_counties = """
SELECT count(*) AS num, County
FROM %s
WHERE Region='绍兴市'
GROUP BY County
ORDER BY count(*) DESC;
"""%data_table

for row in cursor.execute(sql_counties):
    print row[0],row[1]


workbook.Save()
workbook.Close()
excel.Quit()
db.close()