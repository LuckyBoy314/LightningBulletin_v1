# -*- coding: gb2312 -*-
# import arcpy
# from arcpy.sa import *
# # todo.ͳ����ز���
#
# """
#     "Խ��":
# """
# sta = {}
# workspace =ur'D:\bulletinTemp\2015��\������'
# arcpy.env.workspace  = workspace
# infeature = u"D:/Program Files (x86)/LightningBulletin/LightningBulletin.gdb/������"
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
# data = u'D:/bulletinTemp/2015��/data2015��.shp'
# arcpy.FeatureClassToGeodatabase_conversion(data,'/GDB.mdb')

import pyodbc
import win32com.client

#�������ݿ�
#db = pyodbc.connect("DSN=<that Data Source I just created>")
db = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};  DBQ=D:/GDB.mdb;')# Uid=Admin;Pwd=;')
cursor = db.cursor()
data_table = "data2015��"

#��ExcelӦ�ó���
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
#���ļ�����Excel������
workbook = excel.Workbooks.Open(ur'D:\Program Files (x86)\LightningBulletin\Data\�������׵繫��ͼ��.xlsx')


#�㽭�ֵ���ͳ��
sql_regions = """
SELECT count(*) AS num, Region
FROM %s
WHERE Province='�㽭ʡ'
GROUP BY Region
ORDER BY count(*) DESC;
"""%data_table

results = {}
for row in cursor.execute(sql_regions):
    results[row[1]] = row[0]# �� Region��num�����ֵ䣬�������渳ֵ

sheet_regions = workbook.Worksheets(u'ʡ����')
for row in xrange(2,13):
    region = sheet_regions.Cells(row,2).Value
    sheet_regions.Cells(row,1).Value = results[region]


#���˷���ͳ��
sql_counties = """
SELECT count(*) AS num, County
FROM %s
WHERE Region='������'
GROUP BY County
ORDER BY count(*) DESC;
"""%data_table

for row in cursor.execute(sql_counties):
    print row[0],row[1]


workbook.Save()
workbook.Close()
excel.Quit()
db.close()