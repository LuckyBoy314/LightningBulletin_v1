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
import os

year = 2015
cwd = os.getcwd()  # ��ȡ��ǰ����Ŀ¼�����ڳ�����ֲ
# �������ݿ�
db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};',
                             'DBQ=', cwd, '/Data/GDB.mdb;']))  # Uid=Admin;Pwd=;')
cursor = db.cursor()

data_table = ''.join(['data',str(year),'��'])#sql��ѯ��䲻��ʹ������

# ��ExcelӦ�ó���
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
# ���ļ�����Excel������
workbook = excel.Workbooks.Open(''.join([cwd, u'/Data/�������׵繫��ͼ��.xlsx']))

# ************�㽭�ֵ���ͳ��**********
sql = """
SELECT count(*) AS num, Region
FROM %s
WHERE Province='�㽭ʡ'
GROUP BY Region
ORDER BY count(*) DESC
""" % data_table

results = {}
for row in cursor.execute(sql):
    results[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ

sheet = workbook.Worksheets(u'ʡ����')
for row in xrange(2, 13):
    sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

# ********* ���˷���ͳ��***********
sql = """
SELECT count(*) AS num, County
FROM %s
WHERE Region='������'
GROUP BY County
ORDER BY count(*) DESC
""" % data_table

results = {}
for row in cursor.execute(sql):
    results[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ

sheet = workbook.Worksheets(u'�з���')
for row in xrange(3, 9):
    sheet.Cells(row, 3).Value = results[sheet.Cells(row, 2).Value]

# ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
sql = """
SELECT count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,1 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/1/1# And Date_< #YEAR/2/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,2 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/2/1# And Date_< #YEAR/3/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,3 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/3/1# And Date_< #YEAR/4/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,4 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/4/1# And Date_< #YEAR/5/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,5 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/5/1# And Date_< #YEAR/6/1#AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,6 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/6/1# And Date_< #YEAR/7/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,7 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/7/1# And Date_< #YEAR/8/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,8 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/8/1# And Date_< #YEAR/9/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,9 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/9/1# And Date_< #YEAR/10/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,10 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/10/1# And Date_< #YEAR/11/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,11 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/11/1# And Date_< #YEAR/12/1# AND Region='������' and Intensity<0)
union SELECT  count(*) as ��������, -sum(Intensity)/count(*) as ƽ��ǿ�� ,12 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/12/1# And Date_<=#YEAR/12/31# AND Region='������' and Intensity<0)
ORDER BY �·�
""".replace('QUERY_TABLE',data_table).replace('YEAR',str(year))

sheet1 = workbook.Worksheets(u'�·�')
sheet2 = workbook.Worksheets(u'����ǿ�Ȳ���')
i =1#�к�
for row in cursor.execute(sql):
    i+=1
    sheet1.Cells(i,3).Value = row[0]#��������
    sheet2.Cells(i,2).Value = row[1] if row[1] is not None else 0#����ǿ��


# ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
sql = """
SELECT count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,1 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/1/1# And Date_< #YEAR/2/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,2 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/2/1# And Date_< #YEAR/3/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,3 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/3/1# And Date_< #YEAR/4/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,4 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/4/1# And Date_< #YEAR/5/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,5 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/5/1# And Date_< #YEAR/6/1#AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,6 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/6/1# And Date_< #YEAR/7/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,7 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/7/1# And Date_< #YEAR/8/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,8 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/8/1# And Date_< #YEAR/9/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,9 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/9/1# And Date_< #YEAR/10/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,10 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/10/1# And Date_< #YEAR/11/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,11 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/11/1# And Date_< #YEAR/12/1# AND Region='������' and Intensity>0)
union SELECT  count(*) as ��������, sum(Intensity)/count(*) as ƽ��ǿ�� ,12 as �·�
FROM QUERY_TABLE
WHERE (Date_>=#YEAR/12/1# And Date_<=#YEAR/12/31# AND Region='������' and Intensity>0)
ORDER BY �·�
""".replace('QUERY_TABLE',data_table).replace('YEAR',str(year))

i =1#�к�
for row in cursor.execute(sql):
    i+=1
    sheet1.Cells(i,2).Value = row[0] #��������
    sheet2.Cells(i,3).Value = row[1] if row[1] is not None else 0 #����ǿ��

workbook.Save()  # ����EXCEL������
workbook.Close()  # �رչ������ļ�
excel.Quit()  # �ر�EXCELӦ�ó���
db.close()  # �ر���������