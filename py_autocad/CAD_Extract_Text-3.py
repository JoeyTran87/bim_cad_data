from pickle import GLOBAL
from comtypes import IID
import pyautocad
from pyautocad import Autocad, APoint,aDouble, ACAD
import logging,os,time

from win32com.client.build import MakeDefaultArgRepr

import CAD_Set_color_show_hide
from CAD_Set_color_show_hide import *

import pandas as pd
import openpyxl

import re
import win32com.client

import numpy as np
from math import cos, sin
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
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
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
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
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#

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
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
def time_log():
    return time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))

#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#

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

#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
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
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
def main():
    """"""
    global count_text, blocks_dic,texts,ask_att,paths
    time_start = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))    
    # DEBUGGER
    debug_file_name = f"{os.getcwd()}\\py_autocad\\debug\\debug_CAD_Extract_Text.txt" #<---  SỬ DUNG CHO 1 LẦN DEBUG
    # debug_file_name = f"{os.getcwd()}\\debug\\debug-{time.strftime('%y%m%d %H%M%S',time.localtime(time.time()))}.txt" #<---  SỬ DUNG CHO NHIỀU LẦN DEBUG
    logging.basicConfig(filename = debug_file_name,level=logging.INFO, format='%(message)s')
    # logging.disable(logging.CRITICAL) ########### <---  UNCOMMNEND khi không cần debug nữa
    logging.info('-'*20)
    logging.info(f'Program START : {time_start}')
    # BROWSE FOLDER    
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

#--------------------------------------------------------------------------------------------------------#
def open_cad(collect = False):
    """Sử dung win32com để mở Autocad"""
    global acad, cad_doc, model_space,time_, dict_collector,doc_blocks,ms_blocks
    global doc_blocks_names,doc_blocks_contain_att,doc_blocks_contain_att_names,doc_blocks_contain_text,doc_blocks_contain_text_names
    # MỞ AUTOCAD # acad = Autocad(create_if_not_exists=True, visible=True) 
    # check Autocad open
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except:
        acad = win32com.client.Dispatch("AutoCAD.Application")
    # Check Cad idle
    while not acad.GetAcadState().IsQuiescent:
        print("Autocad is busy. Cancel all on going Autocad action and wait in after 5 second")
        time.sleep(5)
        if acad.GetAcadState().IsQuiescent:
            break

    # Autocad DOCUMENTS
    cad_doc_list = list(acad.Documents) 
    # HỎI CHỌN DOCUMENT # SHOW LIST DOC
    [print(f"{cad_doc_list.index(d)} : {d.Name}") for d in cad_doc_list]
    while True:
        try:
            ask_doc = input("\tVui lòng chọn Document: ")    
            if ask_doc == '0' or int(ask_doc):
                cad_doc = cad_doc_list[int(ask_doc)]
                # Model space
                model_space = cad_doc.ModelSpace
                break
        except:
            pass
    # Blocks
    # doc_blocks = [b for b in cad_doc.Blocks if not "*" in b.Name]   #
    get_blocks()
    print("Model_space Elements Count: ",model_space.Count)#sum(1 for x in doc_blocks))#
    print("Model_space Blocks Count: ",len(ms_blocks)) #sum(1 for x in ms_blocks))#

    if collect:
        ms_blocks = [model_space.Item(i) for i in range(model_space.count) if "Block" in model_space.Item(i).ObjectName and not "*" in model_space.Item(i).Name] # ! BAO GỒM CẢ ARRAY (name có startwith *) #get_model_space_blocks() #

def get_blocks():
    global cad_doc,doc_blocks_names ,ms_blocks,count_ms_blocks,count_doc_blocks
    global doc_blocks_contain_att,doc_blocks_contain_att_names,doc_blocks_contain_text,doc_blocks_contain_text_names
    for b in cad_doc.Blocks:
        if not "*" in b.Name:
            count_doc_blocks += 1
            print(b.Name)
            doc_blocks.append(b)
            doc_blocks_names.append(b.Name)            
            try:
                contain_text = False
                for i in range(b.Count):
                    if "Text" in b.Item(i).ObjectName:
                        contain_text = True
                        break
                if contain_text:
                    doc_blocks_contain_text.append(b)
                    doc_blocks_contain_text_names.append(b.Name)
            except Exception as ex:
                print(ex)
                pass
    for i in range(model_space.count):
        b = model_space.Item(i)
        if "Block" in b.ObjectName and not "*" in b.Name:
            ms_blocks.append(b)
            count_ms_blocks +=1
            if b.HasAttributes:
                # doc_blocks_contain_att.append(b)
                doc_blocks_contain_att_names.append(b.Name)


def browse_block(b):
    if not b.ObjectName == "AcDbBlockReference":
        return
    if b.Count > 0:
        for i in range(b.Count):
            if b.Item(i).ObjectName == "AcDbBlockReference":
                browse_block(b.Item(i))
            else:
                pass
def select_region(name = "",promp = None):
    """Chọn đối tượng trên màn hình Autocad"""
    global acad, cad_doc, model_space,time_,dict_collector
    print("\t\tHãy thao tác Lựa chọn trên Autocad")
    select_set = cad_doc.SelectionSets.Add(f"Set_{name}_{time_}")
    if promp == None:
        cad_doc.Utility.Prompt(u"%s\n" % "Chọn đối tượng + Hoàn tất bằng Enter")
    else:
        cad_doc.Utility.Prompt(u"%s\n" % promp)
    select_set.SelectOnScreen()
    print("Số lượng đối tượng chọn: ",select_set.Count)
    
    # Collecting object
    [dict_collector.__setitem__(e.ObjectName,[]) for e in select_set]
    [dict_collector[e.ObjectName].append(e) for e in select_set]

    [print("\t",d," : ",len(dict_collector[d])) for d in dict_collector]
    
    return select_set

def pick_point(promp = None):
    """Pick point Coordinate"""
    global acad, cad_doc, model_space,time_
    if promp == None:
        cad_doc.Utility.Prompt(u"%s\n" % "Chọn điểm để lấy tọa độ")
    else:
        cad_doc.Utility.Prompt(u"%s\n" % promp)
    point_ = cad_doc.Utility.GetPoint()
    print(point_) # [print(i) for i in point_]
    return point_



#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
#--------------------------------------------------------------------------------------------------------#
if __name__ ==  '__main__':
    """"""
    # GLOBAL
    paths = []
    block_names = []
    count_text = 0
    texts = {   "id":[],
                "TextString":[],
                "X":[],
                "Y":[],
                "Z":[]}
    blocks_dic = None
    ask_att = None
    # main()


    acad = None # Autocad Application
    cad_doc = None # Autocad Document
    model_space = None # Autocad Model Space
    time_ = time.strftime("%y%m%d%H%M%S",time.localtime(time.time()))
    dict_collector = {} # CAD COLLECTOR

    # global for doc & ms

    doc_blocks = [] # Document Blocks = Các định nghĩa Block, không phải Cá thể được sử dụng trong Model Space
    doc_blocks_names = []
    doc_blocks_contain_att = []
    doc_blocks_contain_att_names = []
    doc_blocks_contain_text= []
    doc_blocks_contain_text_names = []

    count_doc_blocks = 0
    ms_blocks = [] # Model space blocks (intances) = Các cá thể block được bố trí trong Model space
    count_ms_blocks = 0

    sel_blocks = [] # AcDbBlockReference
    sel_blocks_contain_text = []
    sel_attributes_in_block = []
    sel_texts_in_block = []
    sel_texts = [] #AcDbText
    sel_mtexts = [] #AcDbMText
    sel_mtexts_in_block = []
    sel_attribute = []
    sel_lines = [] #AcDbLine
    sel_polylines = [] #AcDbPolyline

    
    # Open Autcad
    open_cad()   
    
    
    # Selection
    selection_set = select_region()
    
    for e in selection_set:
        obj_name = e.ObjectName
        if obj_name == "AcDbBlockReference":
            sel_blocks.append(e)
            #check block contain attribute
            contain_attribute = e.Name in doc_blocks_contain_att_names
            if contain_attribute:
                pass
            #check Block contain text
            contain_text = e.Name in doc_blocks_contain_text_names
            if contain_text:
                pass

        if obj_name == "AcDbText":
            sel_texts.append(e)
        if obj_name == "AcDbMText": # Multiline Text
            sel_mtexts.append(e)
        if obj_name == "AcDbLine":
            sel_lines.append(e)
        if obj_name == "AcDbPolyline":
            sel_polylines.append(e)

    # Process text
    # for e in selection_set:
    #     if e.ObjectName == "AcDbText":
    #         print(e.InsertionPoint)
    
    # Pick point Coordinate
    # point_ = pick_point()

    