# # -*- coding: utf-8 -*-
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
# #arcpy.env.workspace
# arcpy.CreatePersonalGDB_management("E:/bulletinTemp", "GDB.mdb")
# data = u'E:/bulletinTemp/2015年/data2015年.shp'
# arcpy.FeatureClassToGeodatabase_conversion(data,'E:/bulletinTemp/GDB.mdb')

import pyodbc

db = pyodbc.connect(
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    + r'DBQ=E:/bulletinTemp/GDB.mdb')

cursor1 = db.execute("""
    SELECT [count(*) as 个数, Region]
    FROM [data2015年]
    WHERE [Region] = '绍兴市'
    GROUP BY [Region]""")

while 1:
    row = cursor1.fetchone()
    if not row:
        break
    print row.word
db.close()