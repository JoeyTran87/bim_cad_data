#!/usr/bin/env python
# -*- coding: utf-8 -*-
import openpyxl, json, os
import pandas as pd
import numpy as np

# encoding=utf8  
import sys 
# reload(sys)
# sys.setdefaultencoding('utf8')
import locale
locale.setlocale(locale.LC_ALL, '')



pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")#R:\BimESC\00-BIM STANDARD\PYTHON\ifcopenshell\21.006_Z01_P001_09_.ifc#r"C:\Users\USER\Documents\GitHub\cofico\cofico\FROM BIM MASTER TEMP 210412\Python\pyAutocad\CAD DATA\02B-CL BM-01.txt"


pathCsv = r"C:\Users\USER\Documents\GitHub\cofico\cofico\FROM BIM MASTER TEMP 210412\Python\pyAutocad\CAD DATA\02B-CL BM-01.csv"
pathXls= r"C:\Users\USER\Documents\GitHub\cofico\cofico\FROM BIM MASTER TEMP 210412\Python\pyAutocad\CAD DATA\02B-CL BM-01.xls"

# csvData = pd.read_csv(pathCsv)
# print(csvData) 
# xlsData = pd.read_excel(pathXls)

def txtToDic(pathTxt,dataType,colIndex):
    iniCount = 0
    resCount = 0
    res = []
    with open(pathTxt, "r",encoding="mbcs") as f:
        # print(f.readline().split('\t'))
        # print(len(f.readline().split('\t')))
        keys = f.readline().split('\t')
        lenKey = len(keys)
        for line in f.readlines()[1:]:
            dic = {}
            iniCount +=1
            try:
                atts = line.split('\t')      
                lenAtt = len(atts)
                if lenKey == lenAtt and atts[colIndex] == dataType:
                    for i in range(lenKey):
                        dic[keys[i]] = atts[i]            
                    res.append(dic)
                    resCount +=1
            except :
                pass
    # print (iniCount, resCount)
    # print(len(res[1].split('\t')))
    # print(res[1].split('\t'))
    return res, resCount

if pathTxt is None:
    print("There is no path. Please re-Input (1)")
    pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")
if not os.path.isfile(pathTxt):
    print("File is not exited. Please re-Input (1)")
    pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")

dataType = input("(2) Data Type (Text, Multitext, Line, Polyline, Hatch, Block, Solid, Leader, Circle, Rectangle, Elipse): ")

if dataType is None:
    print("There is no DataType. Please re-Input (2)")
    dataType = input("(2) Data Type (Text, Multitext, Line, Polyline, Hatch, Block, Solid, Leader, Circle, Rectangle, Elipse): ")
if not dataType in ["Text", "Multitext", "Line", "Polyline", "Hatch", "Block", "Solid", "Leader", "Circle", "Rectangle", "Elipse"]:
    print("DataType is not IN LIST. Please re-Input (2)")
    dataType = input("(2) Data Type (Text, Multitext, Line, Polyline, Hatch, Block, Solid, Leader, Circle, Rectangle, Elipse): ")

colIndex = int(input("(3) Data Type Column Index (must be Integer): "))
print(txtToDic(pathTxt,dataType,colIndex))
