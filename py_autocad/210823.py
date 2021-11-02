#!/usr/bin/env python
# coding: utf-8


"""Xử lí toàn bộ Block, Leader, Dim : Color by Block"""
import pyautocad
from pyautocad import Autocad, APoint
import pyautocad, utility
from pyautocad import *
from utility import *
import time, os
from pyautocad import Autocad, APoint,aDouble, ACAD
import unidecode

    
if __name__ ==  '__main__':
    acad = None
    cad_app = None
    cad_doc = None
    path = None
    color_by_block = None
    color_by_layer = None

    acad = Autocad(create_if_not_exists=True, visible=True)
    acad.app
    if acad: print('Autocad Opened')

    # if color_by_block == None and color_by_layer == None:    
    #     color_by_block,color_by_layer = load_color_from_file(acad,r'C:\Users\USER\Documents\GitHub\cofico\cofico\FROM BIM MASTER TEMP 210412\Python\py_autocad\root_source\color.dwg')
    
    # cad_doc = open_file(acad,r'K:\_WFH THU VIEN TRINH BAY DAU THAU\INPUT\15.08.2021.THU MUC CAC BPTC DIEN HINH\_XU LI\Frame W1250xH1530.dwg')#r'K:\_WFH THU VIEN TRINH BAY DAU THAU\INPUT\15.08.2021.THU MUC CAC BPTC DIEN HINH\_XU LI\CT DO BE TONG COT TANG HAM-RAW.dwg')#r'K:\_WFH THU VIEN TRINH BAY DAU THAU\INPUT\15.08.2021.THU MUC CAC BPTC DIEN HINH\_XU LI\test.dwg')

    # color_block_childs_by_block(cad_doc,color_by_block)

    cad_doc = acad.app.Documents[0]

    # block = [cad_doc.Blocks[i] for i in range(cad_doc.Blocks.Count) if cad_doc.Blocks[i].Name == 'Frame W1250xH1530'][0]
    # print(block.Name)
    # print(block.ObjectName)
    # if block.IsDynamicBlock:
    #     block.DynamicBlockReferenceProperty
    
    cad_doc.ModelSpace.Count
    