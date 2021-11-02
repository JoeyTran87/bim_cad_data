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
    cad = None
    cad_app = None
    cad_doc = None
    path = None
    color_by_block = None
    color_by_layer = None

    cad = Autocad(create_if_not_exists=True, visible=True)
    cad.app
    if cad: print('Autocad Opened')
        
        
    print([doc.Name for doc in cad.app.Documents]) 
    cad_doc = cad.app.Documents[0]

    ms = cad_doc.ModelSpace

    dic_items = report_items(cad_doc)

    texts = get_items(cad_doc,in_text = 'Text')

    text_VI(texts)
    