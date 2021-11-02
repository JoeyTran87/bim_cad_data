import json, os, sys
import numpy as np
import matplotlib
import pyautocad
from pyautocad import *
import comtypes.client
import time, unidecode, re

def sayHello():
    print("Hello")
def openCADFile(path):
    try:
        acadApp = comtypes.client.GetActiveObject("AutoCAD.Application")
    except:
        acadApp = comtypes.client.CreateObject("AutoCAD.Application")
    while not acadApp.GetAcadState().IsQuiescent :
        time.sleep(5)
    acadApp.Visible = True
    acad = Autocad()
    if not acad.doc.Name == path.split("\\")[-1]:
        doc = acadApp.Documents.Open(path)
        acad.prompt("Hello, Autocad from Python")    
        print (acad.doc.Name) 

    return acad
class cadExtractor:
    fileName = ""
    def __init__(self):
        self.askPath()
        if len( self.fileName) > 0:
            print( self.fileName)   
    def askPath(self):
        self.fileName = input("File Name Path: ")


def color_layer(acad,cad_doc,layer_color = None):
    acad.prompt("Process Layers")
    cad_layers = cad_doc.Layers
    cad_layers_names = [layer.Name for layer in cad_layers]
    print(cad_layers_names)
    if layer_color == None:
        layer_color = input('Vui lòng chọn mã màu: ')
    for layer in cad_layers:
        layer.Color = layer_color
    cad_doc.Regen (True)


def cad_selector(cad_doc,set_name,promp="Select objects"):
    """Phương thức chọn đối tượng trong file CAD"""
    print(promp)
    cad_doc.Utility.Prompt(u"%s\n" % promp)
    try:
        cad_doc.SelectionSets.Item(set_name).Delete()
    except Exception as ex:
        pass
    selection = cad_doc.SelectionSets.Add(set_name)
    selection.SelectOnScreen()
    return selection


def load_color_from_file(acad,path_color_file = None):
    color_by_block = None
    color_by_layer = None
    cad_doc_color = None
    if path_color_file == None:
        path_color_file = input('Color Path File: ')
    color_file_name = path_color_file.split("\\")[-1]
    docs = acad.Application.Documents # print(cad_doc.Name)
    doc_names = [d.Name for d in docs]    
    if os.path.isfile(path_color_file) and not color_file_name in doc_names:
        acad.app.Documents.Open(path_color_file)
        docs = acad.Application.Documents # print(cad_doc.Name)
        doc_names = [d.Name for d in docs] 
    acad.app.ZoomAll()
    while True:
        for d in docs:
            if d.Name == color_file_name:
                cad_doc_color = d
                break
        if cad_doc_color != None:
            break
        else:
            print('Current files NOT INCLUDE your file')
    entities = cad_doc_color.ModelSpace
    for i in range(entities.Count):
        item = entities[i]
        if 'AcDbMText' in item.ObjectName:
            print(item.TextString)
            if item.TextString == 'Text Color By Block':
                color_by_block = item.TrueColor
            if item.TextString == 'Text Color By Layer':
                color_by_layer = item.TrueColor
    if color_by_block != None and color_by_layer != None:
        cad_doc_color.Close()
    return color_by_block,color_by_layer


def open_file(acad,path = None):
    while True:
        if path == None:
            path = input('Đường dẫn file CAD: ')
        if os.path.isfile(path):
            print('Đã load Path')
            break
        else:
            print('Path không tồn tại')
    file_name = path.split("\\")[-1]
    print(f"File name: {file_name}")

    docs = acad.app.Documents # print(cad_doc.Name)
    doc_names = [d.Name for d in docs]
    if not file_name in doc_names:    
        acad.app.Documents.Open(path)
        docs = acad.app.Documents # print(cad_doc.Name)
        doc_names = [d.Name for d in docs]
    print(doc_names)
    while True:
        for d in docs:
            if d.Name == file_name:
                cad_doc = d
                print('Current files include your file')
                break
        if cad_doc != None:
            break
        else:
            print('Current files NOT INCLUDE your file')
    acad.app.ZoomAll()
    cad_doc.Regen (True)
    return cad_doc


def change_color_to_by_block(cad_doc,name_set,color_by_block):
    """Đổi màu đối tượng thành By Block"""
    elems = cad_doc.SelectionSets.Item(name_set)
    print([e.ObjectName for e in elems])
    for e in elems:
    #     color.SetRGB(0,103,172)
        e.TrueColor = color_by_block
    cad_doc.Regen (True)


def color_block_childs_by_block(cad_doc,color_by_block):
    cad_doc_blocks = cad_doc.Blocks
    cad_doc_blocks_filterer = []
    for block in cad_doc_blocks:
        if not '*' in block.Name:
            for i in range(block.Count):
                try:
                    block.Item(i).TrueColor = color_by_block
                    if 'Dimension' in block.Item(i).ObjectName:                  
                        dim = block.Item(i)                    
                        print(dim.ObjectName)
                        try:
                            dim.DimensionLineColor = color_by_block.EntityColor
                        except Exception as ex:
                            print(ex)
                        try:
                            dim.ExtensionLineColor = color_by_block.EntityColor
                        except Exception as ex:
                            print(ex)
                        try:
                            dim.TextColor = color_by_block.EntityColor
                        except Exception as ex:
                            print(ex)
                    if 'Leader' in block.Item(i).ObjectName:                  
                        leader = block.Item(i)  
                        print(leader.ObjectName)
                        try:
                            leader.DimensionLineColor = color_by_block.EntityColor
                        except Exception as ex:
                            print(ex)
                except Exception as ex:
                    print(ex)
                    pass
            cad_doc_blocks_filterer.append(cad_doc_blocks_filterer)
    cad_doc.Regen (True)

def report_items(cad_doc):
    ms = cad_doc.ModelSpace
    dic = {}
    list_item_names = [ms.Item(i).ObjectName for i in range(ms.Count)]
    for item in list_item_names:
        dic[item] = list_item_names.count(item)
    
    print('\n'.join([f"{item} : {dic[item]}" for item in dic]))
    return dic

def get_items(cad_doc,in_text = 'AcDb'):
    ms = cad_doc.ModelSpace
    list_items = [ms.Item(i) for i in range(ms.Count) if in_text in ms.Item(i).ObjectName]
    return list_items
    
def text_VI(texts):
    for t in texts:
        new_value = t.TextString
        new_value = re.sub("%%C","D= ",unidecode.unidecode(new_value).strip())
        t.TextString = new_value
        
        print(t.TextString)
    