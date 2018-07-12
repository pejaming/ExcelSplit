#!/usr/bin/env python
# -- coding:utf-8 --

import pandas as pd
import os
from enum import Enum
from pandas.core.frame import DataFrame
from os import listdir
from os.path import isfile, join

class Department(Enum):
    android = 'Android技术部'
    ios = "iOS技术部"
    ml = "流量智能部"
    fe = "前端技术部"
    qa = "平台质量效率部"
    pinche = "拼车技术部"
    wuxianserver = "无线后端服务部"
    shejiaoserver = "社交后端服务部"

class Director(Enum):
    android = '刘阳'
    ios = "吕庆春"
    fe = "李丁辉"
    qa = "刘燚"
    pinche = "张大伟"
    wuxianserver = "马鑫"
    shejiaoserver = "刘晓龙"

def SplitKpi(enum,excelName,sheetName,columnName,outputExcelPrefix,outputSheetName):
    data = pd.read_excel(excelName,sheet_name=sheetName,encoding = 'gbk')  #excel文件目录
    for department in enum:
        departmentValue = (str(department.value))
        data1 = data[data[columnName] == departmentValue]
        writer = pd.ExcelWriter(outputExcelPrefix+departmentValue+'.xlsx')
        data1.to_excel(writer, sheet_name=outputSheetName, index=False)
        writer.save()
        writer.close()

def MergeKpi(dir):
    filelist = list_all_files(dir)
    dataAll = DataFrame()
    for file in filelist:
        print('file = ' + file)
        name, ext = os.path.splitext(file)
        if ext != '.xlsx':
            print('ext = ' + ext)
            continue
        print(dir+'/'+file)
        data = pd.read_excel(dir+'/'+file)  # excel文件目录
        if(dataAll.empty):
            dataAll = data

        dataAll = dataAll.append(data,ignore_index=True)
        print(dataAll)


    writer = pd.ExcelWriter(dir+"/all"+ '.xlsx')
    dataAll.to_excel(writer, index=False)
    writer.save()
    writer.close()

    return

##########tools###########################

def list_all_files(file_path):
    return [f for f in listdir(file_path)if isfile(join(file_path, f))]

########## tools end ###########################


#将一个表格按列拆分成n个表格，参数：1.需筛选列值的枚举 2.源文件名 3.需要筛选的sheet名 4.需筛选的列名 5.输出文件的前缀（提升可读性） 6.输出文件的sheet名
SplitKpi(Director,'oneforall/绩效结果收集表模板2018Q2模板-用户增长部-.xlsx','绩效结果收集表模版',"考评人姓名","绩效结果收集表模板2018Q2模板-用户增长部-","名单")
#MergeKpi('allforone')


