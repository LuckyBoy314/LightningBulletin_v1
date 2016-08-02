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
import os
from win32com.client import DispatchEx, constants
from win32com.client.gencache import EnsureDispatch

#����
year = 2015
target_area = u'������'
area = 8256.0
province_area = {u'������':16596.0,u'������':9365.0,u'������':11784.0,u'������':5794.0,u'������':3915.0,
u'������':8256.0,u'����':10919.0,u'̨����':9413.0,u'��ɽ��':1440.0,u'������':8837.0,u'��ˮ��':17298.0}
region_area = {u'Խ����':498.0,u'������':1041.0,u'������':1403.0,u'������':2311.0,u'������':1790.0,u'�²���':1213.0}

#todo ���������Ҫ�����������

cwd = os.getcwd()  # ��ȡ��ǰ����Ŀ¼�����ڳ�����ֲ
# �������ݿ�
db = pyodbc.connect(''.join(['DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};',
                             'DBQ=', cwd, '/Data/GDB.mdb;']))  # Uid=Admin;Pwd=;')
cursor = db.cursor()

data_table = ''.join(['data', str(year), '��'])  # sql��ѯ��䲻��ʹ��Unicode

# ��ExcelӦ�ó���
excel = DispatchEx('Excel.Application')
excel.Visible = False
# ���ļ�����Excel������
workbook = excel.Workbooks.Open(''.join([cwd, u'/Data/����ͼ��ģ��.xlsx']))

EnsureDispatch('Word.Application')
word = DispatchEx('Word.Application')
word.Visible = False
doc = word.Documents.Open(''.join([cwd, u'/Data/����ģ��.docx']))

try:
    # ************�㽭�ֵ���ͳ��**********
    sql = """
    SELECT count(*) AS num, Region
    FROM %s
    WHERE Province='�㽭ʡ'
    GROUP BY Region
    ORDER BY count(*) DESC
    """ % data_table

    #����SQL��ѯ�����˳���¼����������������ȫʡ������
    results = {}
    rank = 0
    for row in cursor.execute(sql):
        results[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ
        rank+=1
        if row [1] == target_area:
            sum_rank_in_province = rank #����������������ȫʡ������

    #��SQL��ѯ���д��Excel
    sheet = workbook.Worksheets(u'ʡ����ͳ��')
    for row in xrange(2, 13):
        sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]

    sum_region = results[target_area]  #��������������
    density_region = sum_region/area  #�����������ܶ�

    #����ȫʡ�����������ܶȺ�ȫʡƽ���ܶ�
    density_province_dict = {}
    density_province = 0
    for key in results:
        density_province_dict[key] = results[key]/province_area[key]
        density_province+= density_province_dict[key]
    density_province/=len(province_area) #ȫʡƽ�������ܶ�
    #�ܶȴӴ�С��������
    density_province_sorted = sorted(density_province_dict.iteritems(),key = lambda d:d[1],reverse=True)
    #���㱾���������ܶ�����
    rank = 0
    for item in density_province_sorted:
        rank+=1
        if item[0] == target_area:
            density_rank_in_province = rank #�����������ܶ���ȫʡ������
            break

    # ********* ���˷���ͳ��***********
    sql = """
    SELECT count(*) AS num, County
    FROM %s
    WHERE Region='������'
    GROUP BY County
    ORDER BY count(*) DESC
    """ % data_table
    #����SQL��ѯ�����˳���¼����������������Сֵ
    results = {}
    rank = 0
    num_region = len(region_area)
    for row in cursor.execute(sql):
        results[row[1]] = row[0]  # �� Region��num�����ֵ䣬�������渳ֵ
        rank+=1
        if rank ==1:
            sum_max_county = row[1]
            sum_max_region = row[0]
        elif rank == num_region:
            sum_min_county = row[1]
            sum_min_region = row[0]
    #SQL��ѯ���д��Excel
    sheet = workbook.Worksheets(u'�з���ͳ��')
    for row in xrange(2, 8):
        sheet.Cells(row, 1).Value = results[sheet.Cells(row, 2).Value]
    #���������Сֵ��ռ�����ı���
    max_region_percent = sum_max_region/float(sum_region)*100
    min_region_percent = sum_min_region/float(sum_region)*100

    #���㱾���������е����ܶ�
    density_region_dict = {}
    for key in results:
        density_region_dict[key] = results[key]/region_area[key]
    #�ܶȴӴ�С��������
    density_region_sorted = sorted(density_region_dict.iteritems(),key = lambda d:d[1],reverse=True)
    #�����С�ܶ�
    density_max_county = density_region_sorted[0][0]
    density_max_region = density_region_sorted[0][1]
    density_min_county = density_region_sorted[num_region-1][0]
    density_min_region = density_region_sorted[num_region-1][1]

    # todo SQL��ѯ�д��Ż�
    # ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
    sql = """
    SELECT count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
    UNION SELECT  count(*) AS ��������, -sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity<0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
    ORDER BY �·�
    """.replace('QUERY_TABLE', data_table).replace('YEAR', str(year))

    sheet = workbook.Worksheets(u'����ͳ��')
    i = 1  # �к�
    month = 0
    sum_month_dict = {}
    negative_intensity_dict= {}
    for row in cursor.execute(sql):
        i += 1
        month+=1
        sheet.Cells(i, 2).Value = row[0]  # ��������
        sheet.Cells(i, 5).Value = negative_intensity_dict[month] = row[1] if row[1] is not None else 0  # ����ǿ��
        sum_month_dict[month] = row[0]

    #����ǿ�ȷ�ֵ�����·�
    negative_intensity_sorted = sorted(negative_intensity_dict.iteritems(),key = lambda d:d[1],reverse=True)
    peak_month_negative_intensity = negative_intensity_sorted[0][0]
    # ************����ͳ�� �µ�����������ƽ��ǿ��(����)**************
    sql = """
    SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,1 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/1/1# AND Date_< #YEAR/2/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,2 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/2/1# AND Date_< #YEAR/3/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,3 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/3/1# AND Date_< #YEAR/4/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,4 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/4/1# AND Date_< #YEAR/5/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,5 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/5/1# AND Date_< #YEAR/6/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,6 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/6/1# AND Date_< #YEAR/7/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,7 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/7/1# AND Date_< #YEAR/8/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,8 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/8/1# AND Date_< #YEAR/9/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,9 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/9/1# AND Date_< #YEAR/10/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,10 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/10/1# AND Date_< #YEAR/11/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,11 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/11/1# AND Date_< #YEAR/12/1#
    UNION SELECT count(*) AS ��������, sum(Intensity)/count(*) AS ƽ��ǿ�� ,12 AS �·�
    FROM QUERY_TABLE
    WHERE Region='������' AND Intensity>=0 AND Date_>=#YEAR/12/1# AND Date_<=#YEAR/12/31#
    ORDER BY �·�
    """.replace('QUERY_TABLE', data_table).replace('YEAR', str(year))
    i = 1  # �к�
    month = 0
    positive_intensity_dict = {}#��¼����ǿ��
    for row in cursor.execute(sql):
        i += 1
        month+=1
        sheet.Cells(i, 3).Value = row[0]  # ��������
        sheet.Cells(i, 6).Value = positive_intensity_dict[month] = row[1] if row[1] is not None else 0  # ����ǿ��
        sum_month_dict[month] += row[0]

    #������ֵ�·�
    positive_intensity_sorted = sorted(positive_intensity_dict.iteritems(),key = lambda d:d[1],reverse=True)
    peak_month_positive_intensity = positive_intensity_sorted[0][0]

    sum_month_sorted = sorted(sum_month_dict.iteritems(),key = lambda d:d[1],reverse=True)
    #�������������·�
    max_month_region = sum_month_sorted[0][0]
    #������������������
    max_months = [sum_month_sorted[0][0],sum_month_sorted[1][0],sum_month_sorted[2][0]]
    max_months.sort()
    #�������������������ռ����
    max_months_percent = 100*(sum_month_sorted[0][1]+sum_month_sorted[1][1]+sum_month_sorted[2][1])/float(sum_region)
    #û�м�⵽�������·�
    months_zero = [i[0] for i in sum_month_sorted if i[1]== 0]
    months_zero.sort()

    sql = """SELECT TOP 1 Date_
    From %s
    Where Region = '%s'
    Order By Date_, OBJECTID
    """%(data_table,target_area.encode('gb2312'))
    for row in cursor.execute(sql):
        first_date = row[0]


    #todo sum_region_lastyear
    sum_region_lastyear = 22122
    day_region = 40
    if 0.05<=(sum_region-sum_region_lastyear)/float(sum_region)<0.1:
        compare_with_lastyear = u'��������'
    elif -0.1<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.05:
        compare_with_lastyear = u'���н���'
    elif 0.1<=(sum_region-sum_region_lastyear)/float(sum_region)<0.5:
        compare_with_lastyear = u'��������'
    elif -0.5<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.1:
        compare_with_lastyear = u'��������'
    elif 0.5<=(sum_region-sum_region_lastyear)/float(sum_region)<0.9:
        compare_with_lastyear = u'�����ϴ�'
    elif -0.9<(sum_region-sum_region_lastyear)/float(sum_region)<=-0.5:
        compare_with_lastyear = u'�����ϴ�'
    elif 0.9<=(sum_region-sum_region_lastyear)/float(sum_region):
        compare_with_lastyear = u'�������'
    elif (sum_region-sum_region_lastyear)/float(sum_region)<=-0.9:
        compare_with_lastyear = u'�������'
    else:
        compare_with_lastyear = u'������ƽ'

    if density_region>density_province:
        compare_with_province = u'����'
    else:
        compare_with_province = u'����'

    p24 = u'%d�����й���������%d�Σ�ƽ�������ܶ�%.2f��/km?��ƽ���ױ���%d�죨����1-1����\
������ĵ���%d����ȣ�%s����ʱ��ֲ�������������Ҫ������%d��%d��%d�£�\
�����µ���ռȫ���ܵ���������%.2f%%���ӿռ�ֲ�������%s��������������࣬%s���١�\
ȫ�е���ƽ���ܶ�%sȫʡƽ����%.2f��/km?����ȫʡ������%s�������ŵ�%dλ��\
����ƽ���ܶ��ŵ�%dλ������1-2����'% (year,sum_region,density_region,day_region,
                                sum_region_lastyear, compare_with_lastyear, max_months[0],max_months[1],max_months[2],
                                max_months_percent,sum_max_county,sum_min_county,compare_with_province,density_province,
                                target_area,sum_rank_in_province,density_rank_in_province)

    rng  = doc.Paragraphs(24).Range
    rng.Text = p24
    rng.InsertParagraphAfter()


    p25 = u'�ݲ���ȫͳ�ƣ�2016��ȫ�����׵��������ֺ���148������Ա�����¹ʡ�\
���ֱ�Ӿ�����ʧ��7788.04��Ԫ����Ӿ�����ʧ677.42��Ԫ��'

    rng  = doc.Paragraphs(25).Range
    rng.Text = p25
    rng.InsertParagraphAfter()

    p109= u'�ӵ���ͳ�������������ֲ���Բ�����%s����������࣬��%d�Σ�%s���٣�ֻ��%d�Σ�\
���߷ֱ�ռȫ���ܵ�������%.2f%%��%.2f%%����ƽ���ܶ�ͳ��������%s�ܶ���ߣ�Ϊ%.2f��/km?��\
%s��ͣ�Ϊ%.2f��/km2������1-1����'%(sum_max_county,sum_max_region,sum_min_county,sum_min_region,
                                            max_region_percent,min_region_percent,
                                            density_max_county,density_max_region,
                                            density_min_county,density_min_region)

    rng  = doc.Paragraphs(109).Range
    rng.Text = p109
    rng.InsertParagraphAfter()

    p110 = u'�ӵ����ܶȿռ�ֲ�ͼ�ϣ���ͼ1-1�����Կ��������������������ݺ����߽�����������ܶȽϸߣ�\
��߳���5��/km2���²������в��ֵ����������ܶȳ���3��/km?��ȫ�д󲿷ֵ��������ܶ�С��2��/km?��'

    rng  = doc.Paragraphs(110).Range
    rng.Text = p110
    rng.InsertParagraphAfter()

    p113= u'���й��ұ�׼�����õ��ױ���ָ�˹��۲⣨��վ��ΧԼ15km�뾶���棩���ױ������Ķ���ƽ����\
������ʡ���綨λ����������㣨��15kmΪ������ֱ�ͳ�Ƹ���15km�뾶��Χ�ڵ��ױ��գ��ٲ�ֵ���㣩��\
2016��ȫ�е����ױ���ƽ��43�죬���Ϊ29�죬���67�졣�ռ�ֲ�������������ƽԭ�����ױ��ս��٣�\
���ϴ󲿺Ͷ��ϲ��������ױ������϶ࣨ��ͼ1-2����'

    rng  = doc.Paragraphs(113).Range
    rng.Text = p113
    rng.InsertParagraphAfter()

    if len(months_zero) ==0:
        months_zero_description = u''
    elif len(months_zero) ==1:
        months_zero_description = u'%d��δ��⵽������'%months_zero[0]
    elif len(months_zero) ==2:
        months_zero_description = u'%d�º�%d��δ��⵽������'%(months_zero[0],months_zero[1])
    else:
        s =  u'�¡�'.join(map(lambda d:str(d),months_zero[:len(months_zero)-1]))
        months_zero_description = u''.join([s,u'�º�%d�¶�δ��⵽������'%months_zero[-1]])
    p115 =u'%d��%s�׵����Ϊ%d��%d�ա��ӷ���ͳ�������������������·ݳ��ֽ�����̬�ֲ�������%s\
����������ֵ������%d�£�%d��%d��%d�����ױ��߷����·ݣ������µ�������ռ������%.2f%%��\
����������ƽ��ǿ�ȵķ�ֵ�ֱ���%d�º�%d�£������·ݲ���ƽ��(��ͼ1-3) ��'%(year,target_area,
                        first_date.month,first_date.day,months_zero_description,
                        max_month_region,max_months[0],max_months[1],max_months[2],max_months_percent,
                        peak_month_positive_intensity,peak_month_negative_intensity)

    rng  = doc.Paragraphs(115).Range
    rng.Text = p115
    rng.InsertParagraphAfter()

    p118 = u'�ӷ�ʱ��ͳ������������������ֵ�����ڵ�18��ʱ�Σ�17:00-18:00����������Ҫ������������㵽���Ͼŵ㣬\
�߸�ʱ���ڵĵ�������ռ������d%������ƽ��ǿ����ʱ��ʲ�״��������������岨������\
����ƽ��ǿ�ȷ�ֵ�ڵ�7��ʱ�Σ�7:00-8:00��������ƽ��ǿ�ȷ�ֵ�ڵ�11��ʱ�Σ�11:00-12:00��(��ͼ1-4)��'

    rng  = doc.Paragraphs(118).Range
    rng.Text = p118
    rng.InsertParagraphAfter()

    p122 = u'������������ǿ�ȷֲ�ͼ�ɼ����������������ǿ�ȳʽ�����̬�ֲ���������������Ҫ������5-60kA�ڣ���ͼ1-5����\
������������������Լռ�ܵ�����87.20%����������Ҫ�ֲ���5-60kA�ڣ���ͼ1-6����\
�������ڸ���������Լռ�ܸ�������91.46%��'
    rng  = doc.Paragraphs(122).Range
    rng.Text = p122
    rng.InsertParagraphAfter()

finally:
    workbook.Save()  # ����EXCEL������
    workbook.Close()  # �رչ������ļ�
    excel.Quit()  # �ر�EXCELӦ�ó���
    doc.Save()
    doc.Close()
    word.Quit()
    db.close()  # �ر���������
