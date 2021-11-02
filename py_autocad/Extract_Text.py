from comtypes import IID
import pyautocad
from pyautocad import Autocad, APoint,aDouble, ACAD
import logging,os,time

import Set_color_show_hide
from Set_color_show_hide import *

def text_in_block(block,item_type = "Text"):
    global count_text, blocks_dic,texts
    for i in range(block.Count):
        try: # xử lí các item trong block
            item = block.Item(i)
            if item_type in item.ObjectName:
                value = item.TextString
                id = item.ObjectId
                location = item.InsertionPoint 
                texts[id] = [value,location]
                count_text += 1
            if 'AcDbBlockReference' in item.ObjectName: # xử lí tiếp các Block con
                sub_block = blocks_dic[item.Name]
                text_in_block(sub_block,item_type = "Text")
        except Exception as ex:
            logging.error(f"{ex}-{time_log()}")
            pass
def time_log():
    return time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))
#-------------------------------------------------------------------#
if __name__ ==  '__main__':
    """"""
    time_start = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))
    # DEBUGGER
    debug_file_name = f"{os.getcwd()}\\debug\\debug2.txt" #<---  SỬ DUNG CHO 1 LẦN DEBUG
    # debug_file_name = f"{os.getcwd()}\\debug\\debug-{time.strftime('%y%m%d %H%M%S',time.localtime(time.time()))}.txt" #<---  SỬ DUNG CHO NHIỀU LẦN DEBUG
    logging.basicConfig(filename = debug_file_name,level=logging.INFO, format='%(message)s')
    # logging.disable(logging.CRITICAL) ########### <---  UNCOMMNEND khi không cần debug nữa
    logging.info('-'*20)
    logging.info(f'Program START : {time_start}')


    # MỞ AUTOCAD
    acad = Autocad(create_if_not_exists=True, visible=True)
    if acad:  logging.info('Autocad Opened')
    
    # MỞ DOCUMENTS AUTOCAD
    cad_doc_list = list(acad.app.Documents)
    
    # SHOW LIST DOC
    [print(f"{i} : {cad_doc_list[i].Name}") for i in range(len(cad_doc_list))]
    
    # HỎI CHỌN DOCUMENT
    while True:
        try:
            ask_doc = input("\tVui lòng chọn Document: ")    
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

    # list_dic_block,list_block_name = create_list_dic_block_2(cad_doc)

    item_type = "Text"
    ms = cad_doc.ModelSpace
    block_names = []
    count_text = 0
    texts = {}

    blocks = cad_doc.Blocks
    blocks_dic = {}
    for i in range(blocks.Count):
        blocks_dic[blocks.Item(i).Name] = blocks.Item(i) 

    for i in range(ms.Count):
        # print(ms.Item(i).ObjectName)
        try: # xử lí các item trong Doc
            item = ms.Item(i)
            if item_type in item.ObjectName:
                count_text += 1
                value = item.TextString
                id = item.ObjectId
                location = item.InsertionPoint 
                texts[id] = [value,location]
            if 'AcDbBlockReference' in item.ObjectName: # lấy ds tên block referent
                block = blocks_dic[item.Name]
                text_in_block(block,item_type = "Text")   
        except Exception as ex:
            logging.error(f"{ex}-{time_log()}")
            pass
    
    print (count_text)
    [print (f"{t} : {texts[t]}") for t in texts]

    time_end = time.strftime("%d-%m-%Y %H:%M:%S",time.localtime(time.time()))
    logging.info(f'Program END : {time_end}')