from comtypes import IID
import pyautocad
from pyautocad import Autocad, APoint,aDouble, ACAD
import logging,os,time

import CAD_Set_color_show_hide
from CAD_Set_color_show_hide import *

import pandas as pd
import openpyxl

import re
import win32com.client

import numpy as np
from math import cos, sin
def process_attributes(item,block_point,block_rot):
    global count_text, blocks_dic,texts,ask_att
    theta = float(block_rot)
    rot = np.array([[cos(theta), -sin(theta)], [sin(theta), cos(theta)]])
    atts = item.GetAttributes()
    for att in atts:
        try:
            att_id = f"{att.ObjectId}{atts.index(att)}"
            att_tag_name = att.TagString
            att_text = att.TextString
            att_point = att.InsertionPoint

            p_matrix = np.array([att_point[0],att_point[1]])
            rotated_p = np.dot(rot,p_matrix)

            # print(f"{att_tag_name}:{att_text}\t{att_point}")
            texts["id"].append(str(att_id))
            texts["TextString"].append(str(att_text))
            texts["X"].append(str(str(round(float(rotated_p[0])+float(block_point[0]),2))))
            texts["Y"].append(str(str(round(float(rotated_p[1])+float(block_point[1]),2))))
            texts["Z"].append(str(str(round(float(att_point[2])+float(block_point[2]),2))))

            count_text += 1
        except:
            pass
def process_text(item,block_point,rot):
    """rot = matrix rotation to product with point matrix"""
    global count_text, blocks_dic,texts
    value = item.TextString
    id = item.ObjectId
    location = item.InsertionPoint 
    p_matrix = np.array([location[0],location[1]])                
    rotated_p = np.dot(rot,p_matrix)    
    # texts[id] = [value,str(round(float(location[0]),2)),str(round(float(location[1]),2)),str(round(float(location[2]),2))]
    texts["id"].append(str(id))
    texts["TextString"].append(str(value))
    texts["X"].append(str(str(round(float(rotated_p[0])+float(block_point[0]),2))))
    texts["Y"].append(str(str(round(float(rotated_p[1])+float(block_point[1]),2))))
    texts["Z"].append(str(str(round(float(block_point[2])+float(block_point[2]),2))))
    count_text += 1

def text_in_block(block,block_point,block_rot,item_type = "Text"):
    global count_text, blocks_dic,texts,ask_att
    # print(block_point)   
    theta = float(block_rot)
    rot = np.array([[cos(theta), -sin(theta)], [sin(theta), cos(theta)]])    

    for i in range(block.Count):
        try: # xử lí các item trong block
            item = block.Item(i)
            if item_type in item.ObjectName:
                process_text(item,block_point,rot)
            if 'AcDbBlockReference' in item.ObjectName: # xử lí tiếp các Block con
                if item.HasAttributes and ask_att == 1:
                    process_attributes(item,block_point,block_rot)                
                sub_block_rot = item.Rotation
                sub_block_point = item.InsertionPoint 
                sub_block_point = [sub_block_point[0],sub_block_point[1],sub_block_point[2]]
                sub_block = blocks_dic[item.Name]
                text_in_block(sub_block,sub_block_point,sub_block_rot,item_type = "Text")
                
        except Exception as ex:
            # print(ex)
            logging.error(f"{ex}-{time_log()}")
            pass
def time_log():
    return time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))
#-------------------------------------------------------------------#


def browser_folder(path_dir,root = ""):
    global paths
    for path in os.listdir(path_dir):
        path_ = f"{path_dir}\\{path}"
        level = len(path_.replace(root,""))- len(path_.replace(root,"").replace("\\",""))
        # print(level)
        if os.path.isdir(path_):            
            # print(f"{' '*2}{path}")
            browser_folder (path_,root = path_)
        elif ".dwg" in path:
            # print(f"\tFile: {path}")
            paths.append(f"{path_dir}\\{path}")

def read_attributes():    
    acad = win32com.client.Dispatch("AutoCAD.Application")
    # iterate through all objects (entities) in the currently opened drawing
    # and if its a BlockReference, display its attributes.
    for entity in acad.ActiveDocument.ModelSpace:
        name = entity.EntityName
        if name == 'AcDbBlockReference':
            HasAttributes = entity.HasAttributes
            if HasAttributes:
                for attrib in entity.GetAttributes():
                    print("{}: {}".format(attrib.TagString, attrib.TextString))

if __name__ ==  '__main__':
    """"""
    time_start = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))    
    # DEBUGGER
    debug_file_name = f"{os.getcwd()}\\py_autocad\\debug\\debug2.txt" #<---  SỬ DUNG CHO 1 LẦN DEBUG
    # debug_file_name = f"{os.getcwd()}\\debug\\debug-{time.strftime('%y%m%d %H%M%S',time.localtime(time.time()))}.txt" #<---  SỬ DUNG CHO NHIỀU LẦN DEBUG
    logging.basicConfig(filename = debug_file_name,level=logging.INFO, format='%(message)s')
    # logging.disable(logging.CRITICAL) ########### <---  UNCOMMNEND khi không cần debug nữa
    logging.info('-'*20)
    logging.info(f'Program START : {time_start}')
    # BROWSE FOLDER
    paths = []
    path_dir =  r"F:\_NGHIEN CUU\_Github\bim_cad_data\py_autocad\dwg"#input("Đường dẫn: ")#r"R:\BimESC\01_PROJECTS\SPAIN WAREHOUSE_EMERGENT\01-INCOME\210720 Full Set Design Submit to ECP"
    browser_folder(path_dir,root = path_dir)
    # [print (p) for p in paths]

    #LOAD DATA CONVERT
    db_path = r"py_autocad\database\Conver_Text_Special.xlsx"
    df_convert = pd.read_excel(db_path,sheet_name="TextSpecial")
    # print(df_convert)

    # MỞ AUTOCAD
    acad = Autocad(create_if_not_exists=True, visible=True)
    if acad:  logging.info('Autocad Opened')
    
    acad2 = win32com.client.Dispatch("AutoCAD.Application")
    # MỞ DOCUMENTS AUTOCAD
    cad_doc_list = list(acad2.Documents)
    
    
    # HỎI CHỌN DOCUMENT
    # SHOW LIST DOC
    [print(f"{i} : {cad_doc_list[i].Name}") for i in range(len(cad_doc_list))]
    while True:
        try:
            ask_doc = input("\tVui lòng chọn Document: ")    
            if ask_doc == '0' or int(ask_doc):
                cad_doc = cad_doc_list[int(ask_doc)]
                break
        except:
            pass
    logging.info(cad_doc.Name)

    # HỎI XỬ LÍ ATTRIBUTE
    processes = ["No","Yes"]
    [print(f"{i} : {processes[i]}") for i in range(len(processes))]
    while True:
        try:
            ask_att = int(input("\tBạn muốn xử lí Block Attribute (y/n): "))
            if ask_att in range(len(processes)):
                break
        except:
            pass

    # XÁC ĐỊNH LOẠI ĐỐI TƯỢNG = TEXT
    item_type = "Text"

    ms = cad_doc.ModelSpace#cad_doc.ModelSpace
    block_names = []
    count_text = 0
    texts = {   "id":[],
                "TextString":[],
                "X":[],
                "Y":[],
                "Z":[]}

    blocks = cad_doc.Blocks
    blocks_dic = {}
    for i in range(blocks.Count):
        blocks_dic[blocks.Item(i).Name] = blocks.Item(i) 

    for item in ms:#i in range(ms.Count):
        # print(ms.Item(i).ObjectName)
        try: # xử lí các item trong Doc
            # item = ms.Item(i)
            item_name =  item.ObjectName
            if item_type in item_name:                
                count_text += 1
                value = item.TextString
                id = item.ObjectId
                location = item.InsertionPoint 
                # texts[id] = [value,str(round(float(location[0]),2)),str(round(float(location[1]),2)),str(round(float(location[2]),2))]
                texts["id"].append(str(id))
                texts["TextString"].append(str(value))
                texts["X"].append(str(str(round(float(location[0]),2))))
                texts["Y"].append(str(str(round(float(location[1]),2))))
                texts["Z"].append(str(str(round(float(location[2]),2))))

            elif 'AcDbBlockReference' in item_name: # lấy ds tên block referent
                block_rot = item.Rotation 
                block_point = item.InsertionPoint                 
                block = blocks_dic[item.Name]

                if ask_att == 1:
                    process_attributes(item,block_point,block_rot)                
                text_in_block(block,block_point,block_rot,item_type = "Text")            
            elif "Dimension" in item_name:
                 pass        
        
        except Exception as ex:
            # print(ex)
            logging.error(f"{ex}-{time_log()}")
            pass
    
    print (f"Text Counter = {count_text}")
    # [print (f"{t} : {texts[t]}") for t in texts]
    time_write = time.strftime("%Y%m%d%H%M%S",time.localtime(time.time()))
    
    excel_path = f"Extract_Text_data{time_write}.xlsx"
    df = pd.DataFrame(texts)
    # print (df)
    with pd.ExcelWriter(excel_path,mode='w') as writter:
        df.to_excel(writter,"Sheet Name",na_rep='NA',startrow=0,startcol=0,engine='openpyxl') 


    pattern = re.compile("\d\d\d\d")

    time_end = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))
    logging.info(f'Program END : {time_end}')