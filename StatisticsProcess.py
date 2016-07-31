# -*- coding: gb2312 -*-
# import arcpy
# from arcpy.sa import *
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
# arcpy.CreatePersonalGDB_management(''.join([cwd,'/data']), "GDB.mdb")
# data = u'D:/bulletinTemp/2015��/data2015��.shp'
# arcpy.FeatureClassToGeodatabase_conversion(data,''.join([cwd,'/data','/GDB.mdb']))

import pyodbc
import win32com.client
import os

year = 2015
cwd = os.getcwd()  # ��ȡ��ǰ����Ŀ¼�����ڳ�����ֲ
# �������ݿ�
db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};',
                             'DBQ=', cwd, '/Data/GDB.mdb;']))  # Uid=Admin;Pwd=;')
cursor = db.cursor()

data_table = ''.join(['data', str(year), '��'])  # sql��ѯ��䲻��ʹ��Unicode

# ��ExcelӦ�ó���
excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = False
# ���ļ�����Excel������
workbook = excel.Workbooks.Open(''.join([cwd, u'/Data/����ͼ��ģ��.xlsx']))

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

sheet = workbook.Worksheets(u'ʡ����ͳ��')
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

sheet = workbook.Worksheets(u'�з���ͳ��')
for row in xrange(2, 8):
    sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

# todo SQL��ѯ�д��Ż�
# ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
sql = """
SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
union SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
ORDER BY �·�
""".replace('QUERY_TABLE', data_table).replace('YEAR', str(year))

sheet = workbook.Worksheets(u'����ͳ��')
i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 2).Value = row[0]  # ��������
    sheet.Cells(i, 5).Value = row[1] if row[1] is not None else 0  # ����ǿ��

# ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
sql = """
SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#AND Region='������'
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
ORDER BY �·�
""".replace('QUERY_TABLE', data_table).replace('YEAR', str(year))

i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # ��������
    sheet.Cells(i, 6).Value = row[1] if row[1] is not None else 0  # ����ǿ��

# ************��ʱ��ͳ�� ʱ�ε���������ʱ��ƽ��ǿ��(����)**************
sql = """
SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,0 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=0
UNION SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=1
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=2
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=3
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=4
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=5
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=6
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=7
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=8
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=9
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=10
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=11
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=12
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,13 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=13
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,14 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=14
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,15 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=15
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,16 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=16
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,17 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=17
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,18 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=18
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,19 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=19
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,20 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=20
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,21 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=21
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,22 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=22
union SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,23 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Val(Time_)=23
ORDER BY ʱ��
""".replace('QUERY_TABLE', data_table)

sheet = workbook.Worksheets(u'��ʱ��ͳ��')
i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 2).Value = row[0]  # ��������
    sheet.Cells(i, 5).Value = row[1] if row[1] is not None else 0  # ����ǿ��

# ************��ʱ��ͳ�� ʱ�ε���������ʱ��ƽ��ǿ��(����)**************
sql = """
SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,0 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=0
UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=1
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=2
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=3
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=4
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=5
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=6
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=7
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=8
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=9
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=10
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=11
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=12
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,13 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=13
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,14 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=14
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,15 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=15
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,16 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=16
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,17 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=17
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,18 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=18
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,19 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=19
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,20 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=20
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,21 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=21
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,22 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=22
union SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,23 AS ʱ��
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Val(Time_)=23
ORDER BY ʱ��
""".replace('QUERY_TABLE', data_table)

i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # ��������
    sheet.Cells(i, 6).Value = row[1] if row[1] is not None else 0  # ����ǿ��

# **********����ǿ�ȷֲ�**************
sql = """
SELECT count(*) AS ��������,0 AS ��߽�,5 AS �ұ߽�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=0 AND Abs(Intensity)<5
union SELECT count(*) AS ��������,5,10
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=5 AND Abs(Intensity)<10
union SELECT count(*) AS ��������,10,15
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=10 AND Abs(Intensity)<15
union SELECT count(*) AS ��������,15,20
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=15 AND Abs(Intensity)<20
union SELECT count(*) AS ��������,20,25
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=20 AND Abs(Intensity)<25
union SELECT count(*) AS ��������,25,30
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=25 AND Abs(Intensity)<30
union SELECT count(*) AS ��������,30,35
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=30 AND Abs(Intensity)<35
union SELECT count(*) AS ��������,35,40
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=35 AND Abs(Intensity)<40
union SELECT count(*) AS ��������,40,45
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=40 AND Abs(Intensity)<45
union SELECT count(*) AS ��������,45,50
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=45 AND Abs(Intensity)<50
union SELECT count(*) AS ��������,50,55
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=50 AND Abs(Intensity)<55
union SELECT count(*) AS ��������,55,60
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=55 AND Abs(Intensity)<60
union SELECT count(*) AS ��������,60,65
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=60 AND Abs(Intensity)<65
union SELECT count(*) AS ��������,65,70
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=65 AND Abs(Intensity)<70
union SELECT count(*) AS ��������,70,75
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=70 AND Abs(Intensity)<75
union SELECT count(*) AS ��������,75,80
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=75 AND Abs(Intensity)<80
union SELECT count(*) AS ��������,80,85
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=80 AND Abs(Intensity)<85
union SELECT count(*) AS ��������,85,90
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=85 AND Abs(Intensity)<90
union SELECT count(*) AS ��������,90,95
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=90 AND Abs(Intensity)<95
union SELECT count(*) AS ��������,95,100
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=95 AND Abs(Intensity)<100
union SELECT count(*) AS ��������,100,150
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=100 AND Abs(Intensity)<150
union SELECT count(*) AS ��������,150,200
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=150 AND Abs(Intensity)<200
union SELECT count(*) AS ��������,200,250
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=200 AND Abs(Intensity)<250
union SELECT count(*) AS ��������,250,300
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=250 AND Abs(Intensity)<300
UNION SELECT count(*) AS ��������,300,1000
FROM QUERY_TABLE
WHERE Region='������' AND Intensity<0 AND Abs(Intensity)>=300
ORDER BY ��߽�
""".replace("QUERY_TABLE", data_table)

sheet = workbook.Worksheets(u'ǿ�ȷֲ�ͳ��')
i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 3).Value = row[0]  # ��������

# ***********����ǿ�ȷֲ�************
sql = """
SELECT count(*) AS ��������,0 AS ��߽�,5 AS �ұ߽�
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=0 AND Intensity<5
union SELECT count(*) AS ��������,5,10
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=5 AND Intensity<10
union SELECT count(*) AS ��������,10,15
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=10 AND Intensity<15
union SELECT count(*) AS ��������,15,20
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=15 AND Intensity<20
union SELECT count(*) AS ��������,20,25
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=20 AND Intensity<25
union SELECT count(*) AS ��������,25,30
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=25 AND Intensity<30
union SELECT count(*) AS ��������,30,35
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=30 AND Intensity<35
union SELECT count(*) AS ��������,35,40
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=35 AND Intensity<40
union SELECT count(*) AS ��������,40,45
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=40 AND Intensity<45
union SELECT count(*) AS ��������,45,50
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=45 AND Intensity<50
union SELECT count(*) AS ��������,50,55
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=50 AND Intensity<55
union SELECT count(*) AS ��������,55,60
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=55 AND Intensity<60
union SELECT count(*) AS ��������,60,65
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=60 AND Intensity<65
union SELECT count(*) AS ��������,65,70
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=65 AND Intensity<70
union SELECT count(*) AS ��������,70,75
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=70 AND Intensity<75
union SELECT count(*) AS ��������,75,80
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=75 AND Intensity<80
union SELECT count(*) AS ��������,80,85
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=80 AND Intensity<85
union SELECT count(*) AS ��������,85,90
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=85 AND Intensity<90
union SELECT count(*) AS ��������,90,95
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=90 AND Intensity<95
union SELECT count(*) AS ��������,95,100
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=95 AND Intensity<100
union SELECT count(*) AS ��������,100,150
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=100 AND Intensity<150
union SELECT count(*) AS ��������,150,200
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=150 AND Intensity<200
union SELECT count(*) AS ��������,200,250
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=200 AND Intensity<250
union SELECT count(*) AS ��������,250,300
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=250 AND Intensity<300
UNION SELECT count(*) AS ��������,300,1000
FROM QUERY_TABLE
WHERE Region='������' AND Intensity>=300
ORDER BY ��߽�
""".replace("QUERY_TABLE", data_table)

i = 1  # �к�
for row in cursor.execute(sql):
    i += 1
    sheet.Cells(i, 4).Value = row[0]  # ��������

workbook.Save()  # ����EXCEL������
workbook.Close()  # �رչ������ļ�
excel.Quit()  # �ر�EXCELӦ�ó���
db.close()  # �ر���������