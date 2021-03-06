# -*- coding: utf-8 -*-
import math
import os
import time
import arcpy
from arcpy.sa import *


# 获取当前工作区域的范围
def getExtents(infeature):
    extents = [row[0].extent for row in arcpy.da.SearchCursor(infeature, "SHAPE@")]
    xmax = max([e.XMax for e in extents])
    xmin = min([e.XMin for e in extents])
    ymax = max([e.YMax for e in extents])
    ymin = min([e.YMin for e in extents])
    return (xmax, xmin, ymax, ymin)


# 创建网格，根据工作区域范围和网格大小创建网格
def makeGrids(extents, cell_size):
    xmax, xmin, ymax, ymin = extents
    # 计算经纬方向一公里的经纬距离
    midLat = (ymin + ymax) / 2 * (math.pi / 180)
    X1Km = 1 / (math.cos(midLat) * 111)
    Y1Km = 1.0 / 111

    cell_size = int(cell_size)
    # 网格大小
    XWidth = X1Km * cell_size
    YHeight = Y1Km * cell_size

    # 范围
    Left = xmin - XWidth
    Right = xmax + XWidth
    Bottom = ymin - YHeight
    Top = ymax + YHeight

    # Set the extent of the fishnet
    originCoordinate = ''.join([str(Left) , " " , str(Bottom)])
    oppositeCoorner = ''.join([str(Right) , " " ,str(Top)])

    # Set the orientation
    yAxisCoordinate = ''.join([str(Left) , " " , str(Bottom + 1)])

    # Enter 0 for width and height - these values will be calculated by the tool
    cellSizeWidth = str(XWidth)
    cellSizeHeight = str(YHeight)

    outputGrids = ''.join(["GRID" ,str(cell_size)])
    arcpy.CreateFishnet_management(outputGrids, originCoordinate, yAxisCoordinate,
                                   cellSizeWidth, cellSizeHeight, '0', '0', oppositeCoorner, 'NO_LABELS', '#',
                                   'POLYGON')

    return outputGrids


def clipData(origin_data, grid_feature, cell_size):
    output_clip = ''.join(["clip" , str(cell_size)])
    arcpy.Clip_analysis(origin_data, grid_feature, output_clip)
    return output_clip


def lightningDensity(clip_feature, grid_feature, cell_size):
    #if arcpy.Exists("lightningDensity.tif"):
    #   return
    intersect_feature = ''.join(["intersect" , str(cell_size)])
    arcpy.Intersect_analysis([clip_feature, grid_feature], intersect_feature)

    frequency_table = ''.join(["intersect" , str(cell_size) , "_frequency"])
    field_frequency = ''.join(["FID_GRID" ,str(cell_size)])
    arcpy.Frequency_analysis(intersect_feature, frequency_table, field_frequency)

    layer = ''.join(["GRID" ,str(cell_size)])
    arcpy.MakeFeatureLayer_management(grid_feature, layer)

    field_join = ''.join(["FID_GRID" , str(cell_size)])
    arcpy.AddJoin_management(layer, "OID", frequency_table, field_join)

    density_points = "densityPoints"
    arcpy.FeatureToPoint_management(layer, density_points)
    lightning_density_raster = Spline(density_points, "FREQUENCY")

    lightning_density_raster.save("lightningDensity")


def lightningDay(clip_feature, grid_feature, cell_size):
    #if arcpy.Exists("lightningDay.tif"):
    #    return

    # 建立圆形缓冲区
    point_feature = ''.join(["point" , str(cell_size)])
    arcpy.FeatureToPoint_management(grid_feature, point_feature)

    buffer_feature = ''.join(["buffer" , str(cell_size)])
    buffer_distance = ''.join([str(cell_size) ," " , "Kilometers"])
    arcpy.Buffer_analysis(point_feature, buffer_feature, buffer_distance)

    intersect_feature =''.join(["intersect" , str(cell_size)])
    arcpy.Intersect_analysis([clip_feature, buffer_feature], intersect_feature)

    field_frequency = ''.join(["FID_buffer",cell_size])
    frequency_table1 = ''.join(["intersect" , str(cell_size) , "_frequency1"])
    arcpy.Frequency_analysis(intersect_feature, frequency_table1, ["Date", field_frequency])

    frequency_table2 = ''.join(["intersect" , str(cell_size) ,"_frequency2"])
    arcpy.Frequency_analysis(frequency_table1, frequency_table2, field_frequency)

    layer = ''.join(["point" , str(cell_size)])
    arcpy.MakeFeatureLayer_management(point_feature, layer)

    field_join = ''.join(["FID_buffer",cell_size])
    arcpy.AddJoin_management(layer, "ORIG_FID", frequency_table2, field_join)

    day_points = "dayPoints"
    arcpy.FeatureToPoint_management(layer, day_points)
    lightning_day_raster = Spline(day_points, "FREQUENCY")
    lightning_day_raster.save("lightningDay")


def geoProcess(target_area, origin_data,  density_cell="10", day_cell="15"):
    #if arcpy.Exists("lightningDay") and arcpy.Exists("lightningDensity"):
    #   return
    cwd = os.getcwd()
    infeature = ''.join([cwd,u'/data/LightningBulletin.gdb/',target_area])
    # 获取当前extent
    extents = getExtents(infeature)
    ####制作电闪密度插值点文件#######
    # 制作网格文件
    grid_feature = makeGrids(extents, density_cell)
    # 剪裁数据
    clip_feature = clipData(origin_data, grid_feature, density_cell)
    # 制作电闪密度插值点
    lightningDensity(clip_feature, grid_feature, density_cell)

    # 制作掩盖文件
    mask_feature = "densityMask"
    arcpy.Erase_analysis(grid_feature, infeature, mask_feature)

    ####制作电闪雷暴日插值点文件####
    # 制作网格文件
    grid_feature = makeGrids(extents, day_cell)
    # 剪裁数据
    clip_feature = clipData(origin_data, grid_feature, day_cell)
    # 制作电闪雷暴日插值点
    lightningDay(clip_feature, grid_feature, day_cell)
    # 制作掩盖文件
    mask_feature = "dayMask"
    arcpy.Erase_analysis(grid_feature, infeature, mask_feature)


if __name__ == "__main__":
    datetime = u"2015年"
    target_area = u"绍兴市"
    cwd = os.getcwd()

    #todo 新建数据库占用了十几秒的时间，可以考虑在之前并行处理
    workpath = ''.join([cwd,u"/bulletinTemp/", datetime])
    workspace = ''.join([workpath,'/',target_area,'.gdb'])
    if not arcpy.Exists(workspace):
        arcpy.CreateFileGDB_management(workpath,''.join([target_area,'.gdb']))
    arcpy.env.workspace = workspace
    arcpy.env.outputCoordinateSystem = arcpy.SpatialReference("WGS 1984")
    arcpy.env.overwriteOutput = True
    # 输入数据

    database = ''.join([cwd,u'/bulletinTemp/',datetime,u'/GDB.gdb/data',datetime])
    start = time.clock()
    # ***********************测试程序*********************************"
    geoProcess(target_area, database)
    # ***********************测试程序*********************************"
    end = time.clock()
    elapsed = end - start
    print "Time used: %.6fs, %.6fms" % (elapsed, elapsed * 1000)
