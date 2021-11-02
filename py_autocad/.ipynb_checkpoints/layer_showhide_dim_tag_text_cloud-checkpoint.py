#!/usr/bin/env python
# coding: utf-8

import pyautocad
from pyautocad import Autocad, APoint,aDouble, ACAD
import logging,os,time


if __name__ ==  '__main__':
    """"""
    # DEBUGGER
    debug_file_name = f"{os.getcwd()}\\debug\\debug.txt" #<---  SỬ DUNG CHO 1 LẦN DEBUG
    # debug_file_name = f"{os.getcwd()}\\debug\\debug-{time.strftime('%y%m%d %H%M%S',time.localtime(time.time()))}.txt" #<---  SỬ DUNG CHO NHIỀU LẦN DEBUG
    logging.basicConfig(filename = debug_file_name,level=logging.INFO, format='%(message)s')
    # logging.disable(logging.CRITICAL) ########### <---  UNCOMMNEND khi không cần debug nữa
    logging.info('Program START')
    
    # MỞ AUTOCAD
    acad = Autocad(create_if_not_exists=True, visible=True)
    acad.app
    if acad:  logging.info('Autocad Opened')
    
    # MỞ DOCUMENTS AUTOCAD
    cad_doc_list = list(acad.app.Documents)
    
    # SHOW LIST DOC
    [print(f"{i} : {cad_doc_list[i].Name}") for i in range(len(cad_doc_list))]
    
    # HỎI CHỌN DOCUMENT
    while True:
        try:
            ask_doc = input("Vui lòng chọn Document: ")    
            if ask_doc == '0' or int(ask_doc):
                cad_doc = cad_doc_list[int(ask_doc)]
                break
        except:
            continue
    logging.info(cad_doc.Name)

    # MODEL SPACE
    model_space = cad_doc.ModelSpace
    paper_space = cad_doc.PaperSpace
    logging.info([model_space.Name,paper_space.name])

    # LAYOUT
    layouts = list(cad_doc.Layouts)
    [print(f"{i} : {layouts[i].Name}") for i in range(len(layouts))]    
    logging.info([layout.Name for layout in layouts])
    # HỎI CHỌN LAYOUT
    while True:
        try:
            ask_layout = input("Vui lòng chọn Layout: ")    
            if ask_layout == '0' or int(ask_layout):
                layout = layouts[int(ask_layout)]
                break
        except:
            continue
    logging.info(layout.Name)

    # VIEWPORTS
    # viewports = list(cad_doc.Viewports)
    # [print(f"{i} : {viewports[i].Name}") for i in range(len(viewports))]    
    # logging.info([viewport.Name for viewport in viewports])
    # # HỎI CHỌN Viewport
    # while True:
    #     try:
    #         ask_viewport = input("Vui lòng chọn Viewport: ")    
    #         if ask_viewport == '0' or int(ask_viewport):
    #             viewport = viewports[int(ask_viewport)]
    #             break
    #     except:
    #         continue
    # logging.info(viewport.Name)
    print("Bạn cần active vào Viewport bạn muốn xử lí")
    active_viewport = cad_doc.ActivePViewport
    logging.info(f"{active_viewport.ObjectName} {round(active_viewport.Width)} x {round(active_viewport.Height)}")
    
    #  VIEWPORT LAYER
    vp_doc = active_viewport.Document
    logging.info(f"{vp_doc.Name}")


    # layers = list(active_viewport.Layer)
    # [print(layer.Name) for layer in layers]
