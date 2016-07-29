# -*- coding: utf-8 -*-
import DataPreprocess
import GeoProcess
import Mapping
import os
import arcpy
import sys


"""
主要输入：
    origin_data_path:原始数据路径：
    target_area：目标区域
    datetime:时间，年度公报 月报 日报 任意时间间隔
    density_cell, density_class:电闪密度网格大小，图例分类个数
    day_cell, day_class：地闪雷暴日网格大小，图例分类个数
    out_type:输出格式
    out_path:输出路径
主要输出：
    图片(相应格式）
    统计信息
"""

#todo 字符串连接需要优化
def mainProcess(origin_data_path, datetime, target_area, density_cell, density_class,
                day_cell, day_class, out_type, out_path=None):

    database = DataPreprocess.preProcess(origin_data_path, datetime)
    workspace = u"D:/bulletinTemp/" + datetime + u"/" + target_area
    arcpy.env.overwriteOutput = True
    # 设置环境
    if not os.path.exists(workspace):
        os.makedirs(workspace)
    arcpy.env.workspace = workspace
    arcpy.env.outputCoordinateSystem = arcpy.SpatialReference("WGS 1984")

    # 地理处理，生成作图文件
    GeoProcess.geoProcess(target_area, database, density_cell, day_cell)

    # 地图处理
    """
    暂时采用默认参数
    density_class = 10
    day_class = 10
    out_type =
    out_path =
    """
    out_path = u"D:/bulletinTemp/" + datetime
    Mapping.mappingProcess(target_area, datetime,out_path=out_path)


if __name__ == "__main__":

    try:
        parms = sys.argv[1].decode("utf-8")
        #kws = eval(parms)
        kws = ast.literal_eval(sys.argv[1])
        mainProcess(**kws)
    except Exception,inst:
        text_table = u'D:/ceshi.txt'
        with open(text_table, 'w') as out_f:
            out_f.write(unicode(chardet.detect(sys.argv[1])))
            out_f.write('\n')
            out_f.write(u"The error messages:\n")
            out_f.write(unicode(type(inst)))    # the exception instance
            out_f.write(u'\n')
            out_f.write(unicode(inst.args))     # arguments stored in .args
            out_f.write(u'\n')
            out_f.write(unicode(inst))

    # datetime = u"2012年"
    # # 数据预处理。产生某一时间datetime的基础数据，供可区域使用，预处理之后再接着下面的处理
    # origin_data_path = DataPreprocess.dialogOpenFile()
    #
    # start = time.clock()
    # # ***********************测试程序*********************************"
    # database = DataPreprocess.preProcess(origin_data_path, datetime)  # 生成基础数据
    #
    # # 根据目标区域和时间生成临时文件夹，作为工作空间
    # targets = [u"绍兴市", u"柯桥区", u"上虞区", u"诸暨市", u"嵊州市", u"新昌县"]
    # for target_area in targets:
    #     workspace = u"D:/bulletinTemp/" + datetime + u"/" + target_area
    #     arcpy.env.overwriteOutput = True
    #     # 设置环境
    #     if not os.path.exists(workspace):
    #         os.makedirs(workspace)
    #     arcpy.env.workspace = workspace
    #     arcpy.env.outputCoordinateSystem = arcpy.SpatialReference("WGS 1984")
    #
    #     # 地理处理，生成作图文件
    #     density_cell = "10"
    #     day_cell = "15"
    #     GeoProcess.geoProcess(target_area, database, density_cell, day_cell)
    #
    #     # 地图处理
    #     """
    #     暂时采用默认参数
    #     density_class = 10
    #     day_class = 10
    #     out_type =
    #     out_path =
    #     """
    #     out_path = u"D:/bulletinTemp/" + datetime
    #     Mapping.mappingProcess(target_area, datetime,
    #                            out_path=out_path)  # ,density_class, day_class, out_path, out_type)
    # # ***********************测试程序*********************************"
    # end = time.clock()
    # elapsed = end - start
    # print("Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000))
