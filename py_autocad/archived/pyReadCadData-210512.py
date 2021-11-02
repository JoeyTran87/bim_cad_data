import openpyxl, json, os
import pandas as pd
import numpy as np

import sys 
import locale
locale.setlocale(locale.LC_ALL, '')




def txtToDic(pathTxt,dataType,colIndex,targetProperties,region,level,translateCs,category,project,disciptline):
    iniCount = 0
    resCount = 0
    res = []
    targetCol = []

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
                dic["ProjectNumber"] = project
                dic["Disciptline"] = disciptline
                dic["TranslateCoordinateSystem"] = translateCs
                dic["BuildingRegion"] = region
                dic["BuildingLevel"] = level
                dic["DataType"] = dataType
                dic["Category"] = category
                atts = line.split('\t')      
                lenAtt = len(atts)
                if lenKey == lenAtt and atts[int(colIndex)] == dataType:
                    for cii in targetCol:
                        dic[keys[cii]] = atts[cii]
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


pathTxt =dataType=colIndex=targetColIndices=declareRegion=declareLevel=translateCs=category=project=disciptline=askAuto = ""

pathTxt = input("""START-----------START----------START------------START----------START
Step (1) File Path of TXT file (include file name with extension .txt): """)
askAuto = input("Do you want Auto Generate (Yes/No):")
if askAuto == "No":
    dataType = input("""---------------------
    Step (2) Data Type 
    (with no Input, DEFAULT will be applied)
    Ex: Text, Multitext, Line, Polyline
    , Hatch, Block, Solid, Leader
    , Circle, Rectangle, Elipse
    Type here: """)

    colIndex = input("""----------------------
    (3) Column Index of Data Type (must be Integer)
    (with no Input, DEFAULT will be applied)
    Type here:""")

    targetColIndices = input("""---------------------
    (4) Properties of Targeted Data 
    (with no Input, DEFAULT will be applied)
    (Must be list of Name)
    Ex: Layer,Position X,Position Y,Position Z,Rotation,Value: """)

    declareRegion= input("""----------------------
    (5) BUILDInG REGION Abbreviation of Data:
    (with no Input, DEFAULT will be applied)
    Ex:
    BA (Basement), 
    PD (Podium),
    LS (Landscape),
    A (Tower A),
    B (Tower B)
    Type here: """)

    declareLevel = input("""----------------------
    (6) Building LEVEL Abbreviation of Data:
    (with no Input, DEFAULT will be applied)
    Ex: 02B, 01B, GF, 01F 01FM
    Type here: """)

    translateCs = input("""----------------------
    (7) Data Coordianate Translate X & Y & Z & Rotation at New Coordinate:
    (with no Input, DEFAULT will be applied)
    (Should be distance measured to translate CAD Data Coordiate to BIM model Coordinate
    and rotation angle at BIM Model Coordinate
    Unit is Milimeter and Degree, seperator is 1 space)
    Ex: 375 65553 0 0
    Type here: """)

    category = input("""----------------------
    (8) Information Code 
    (with no Input, DEFAULT will be applied)
    Ex: 
    CL (column)
    BM (Beam)
    FD (Foundation)
    Type here: """)




ext = ".txt"
nameWithNotExt = pathTxt.split(".txt")[0].split("\\")[-1]
fileDir = pathTxt.split(nameWithNotExt+ext)[0]


nameCodes = nameWithNotExt.split("_")

if len(dataType) == 0:
    dataType = "Text"
if len(colIndex) == 0:
    colIndex = "1"
if len(targetColIndices) == 0:
    targetColIndices = "Layer,Position X,Position Y,Position Z,Rotation,Value"
if len(project) == 0:
    project = nameCodes[0]
if len(disciptline) == 0:
    disciptline = nameCodes[1]
if len(declareRegion) == 0:
    declareRegion = nameCodes[2]
if len(declareLevel) == 0:
    declareLevel = nameCodes[3]
if len(category) == 0:
    category = nameCodes[4]
if len(translateCs) == 0:
    translateCs = nameCodes[5]   
    
jsonPath = fileDir+nameWithNotExt+"-"+category+"-"+dataType+".json"

data = txtToDic(pathTxt,dataType,colIndex,targetColIndices,declareRegion,declareLevel,translateCs,category,project,disciptline)

writeJson = writeTxtToJson(jsonPath,data)
print("SUCCEEDED TRANSFER JSON FILE :",jsonPath)
print("""END----------END------------END------------END----------END""")