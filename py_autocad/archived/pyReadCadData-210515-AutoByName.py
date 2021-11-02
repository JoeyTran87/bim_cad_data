import openpyxl, json, os
import pandas as pd
import numpy as np

import sys 
import locale
locale.setlocale(locale.LC_ALL, '')

pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")


def txtToDic(pathTxt,dataType,colIndex,targetProperties,region,level,translateCs,category):
    iniCount = 0
    resCount = 0
    res = []
    targetCol = []
    # listFilter = []
    # try:
    #     filters = filterer.split("----")
    #     if len(filters) > 1:
    #         for f in filters:
    #             listFilter.append(f.split("--"))
    #             # keyFilter = f.split("--")[0]
    #             # operateFilter = f.split("--")[1]
    #             # valueFilter = f.split("--")[2]
        
    #     elif len(filters) == 1:
    #          listFilter.append(filters.split("--"))
    #         # keyFilter = filters.split("--")[0]
    #         # operateFilter = filters.split("--")[1]
    #         # valueFilter = filters.split("--")[2]
    # except :
    #     pass

    with open(pathTxt, "r",encoding="mbcs") as f:
        keys = f.readline().split('\t')
        targets = targetProperties.split(",")
        for ii in range(len(keys)):
            for prop in targets:
                if prop ==  keys[ii]:
                    targetCol.append(ii)            
        lenKey = len(keys)
        for line in f.readlines()[1:]:
            dic = {}
            iniCount +=1
            try:
                dic["TranslateCoordinateSystem"] = translateCs
                dic["BuildingRegion"] = region
                dic["BuildingLevel"] = level
                dic["DataType"] = dataType
                dic["Category"] = category
                atts = line.split('\t')      
                lenAtt = len(atts)
                if lenKey == lenAtt and atts[colIndex] == dataType:
                    for cii in targetCol:
                        dic[keys[cii]] = atts[cii]
                        
                # newdic = filteringDic(dic,listFilter)
                    res.append(dic)
                    resCount +=1
            except :
                pass
    return res
def filteringDic(dic,listFilter):
    newdic = {}
    for d in dic:
        for f in listFilter:
            if f[1].lower() == "contain":
                if f[0] in dic.get(d):
                    pass
                pass
            if f[1].lower() == "startwith":
                pass
def writeTxtToJson(path,data):
	with open(path,'w',encoding="mbcs") as f:				
		f.write(json.dumps(data,indent=4))
	return "Succeeded"

#INPUT
# if pathTxt is None:
#     print("There is no path. Please re-Input (1)")
#     pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")
# if not os.path.isfile(pathTxt):
#     print("File is not exited. Please re-Input (1)")
#     pathTxt = input("(1) Path of TXT file (include file name with extension .txt): ")

dataType = input("""(2) Data Type 
(Text, Multitext, Line, Polyline
, Hatch, Block, Solid, Leader
, Circle, Rectangle, Elipse)
----------------------
Type here: """)

# if dataType is None:
#     print("There is no DataType. Please re-Input (2)")
#     dataType = input("(2) Data Type (Text, Multitext, Line, Polyline, Hatch, Block, Solid, Leader, Circle, Rectangle, Elipse): ")
# if not dataType in ["Text", "Multitext", "Line", "Polyline", "Hatch", "Block", "Solid", "Leader", "Circle", "Rectangle", "Elipse"]:
#     print("DataType is not IN LIST. Please re-Input (2)")
#     dataType = input("(2) Data Type (Text, Multitext, Line, Polyline, Hatch, Block, Solid, Leader, Circle, Rectangle, Elipse): ")

colIndex = int(input("(3) Column Index of Data Type (must be Integer): "))

targetColIndices = input("""(4) Properties of Targeted Data 
---Must be list of Name---
Ex: Layer,Position X,Position Y,Position Z,Rotation,Value: """)#35,36,37,40#Layer,Position X,Position Y,Position Z,Rotation,Value

declareRegion= input("""(5) BUILDInG REGION Abbreviation of Data:
Ex: 
BA (Basement), 
PD (Podium),
LS (Landscape),
A (Tower A),
B (Tower B)
----------------------
Type here: """)

declareLevel = input("""(6) Building LEVEL Abbreviation of Data:
(Ex: 02B, 01B, GF, 01F 01FM)
----------------------
Type here: """)
# filterer = input("""(7) Filterer:
# Must be List of Couple Condition like cases of syntax 
# ColumnName--operater--FilterValue
# (Ex: 
# Case1: Value--contain--VH
# Case2: Value--contain--VH----Value--contain--DH
# Case3: Value--startwith--VH
# Case2: Value--startwith--VH----Value--startwith--DH)
# ----------------------
# Type here: """)

translateCs = input("""(7) Data Coordianate Translate X & Y & Z & Rotation at New Coordinate:
---Should be distance measured to translate CAD Data Coordiate to BIM model Coordinate
and rotation angle at BIM Model Coordinate
Unit is Milimeter and Degree, seperator is 1 space---
Ex:
375 65553 0 0
----------------------
Type here: """)

category = input("""(8) Information Code
Ex: 
CL (column)
BM (Beam)
FD (Foundation)
----------------------
Type here: """)


data = txtToDic(pathTxt,dataType,colIndex,targetColIndices,declareRegion,declareLevel,translateCs,category)

ext = ".txt"
nameWithNotExt = pathTxt.split(".txt")[0].split("\\")[-1]
fileDir = pathTxt.split(nameWithNotExt+ext)[0]
jsonPath = fileDir+nameWithNotExt+"-"+category+"-"+dataType+".json"

writeJson = writeTxtToJson(jsonPath,data)
print("SUCCEEDED TRANSFER JSON FILE :",jsonPath)