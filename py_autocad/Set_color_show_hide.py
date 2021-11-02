#!/usr/bin/env python
# coding: utf-8

from comtypes import IID
import pyautocad
from pyautocad import Autocad, APoint,aDouble, ACAD
import logging,os,time

from functools import lru_cache
from timeit import repeat

def cad_selector(cad_doc,set_name,promp="Select objects"):
    """Chọn đối tượng trong file CAD"""
    print(promp)
    cad_doc.Utility.Prompt(u"%s\n" % promp)
    try:
        cad_doc.SelectionSets.Item(set_name).Delete()
    except Exception as ex:
        pass
    selection = cad_doc.SelectionSets.Add(set_name)
    selection.SelectOnScreen()
    return selection

def get_items(cad_doc,in_text = 'AcDb',in_xref = False,in_block = False):
    """Chọn các đối tượng mà ObjectName có chứa in_text"""
    ms = cad_doc.ModelSpace
    list_items = [ms.Item(i) for i in range(ms.Count) if in_text in ms.Item(i).ObjectName]
    return list_items
    
def open_file(acad,path = None):    
    """Mở file Autocad"""
    if path == None:
        while True:
            path = input('Đường dẫn file CAD: ')
            if os.path.isfile(path):
                print('Đã load Path')
                break
            else:
                print('Path không tồn tại')    
    file_name = path.split("\\")[-1]
    # print(f"File name: {file_name}")
    docs = acad.app.Documents # print(cad_doc.Name)
    doc_names = [d.Name for d in docs]
    while True:
        if not file_name in doc_names:    
            acad.app.Documents.Open(path)
            docs = acad.app.Documents # print(cad_doc.Name)
            doc_names = [d.Name for d in docs]    
        for d in docs:
            if d.Name == file_name:
                cad_doc = d
                print('Thành công mở file xref')
                break
        if cad_doc != None:
            break
        else:
            print('Không thể mở file xref')
    acad.app.ZoomAll()
    cad_doc.Regen (True)
    return cad_doc
def show_hide_items_in_doc(doc,item_type = "Text",show =False):
    ms = doc.ModelSpace
    block_names = []
    for i in range(ms.Count):
        try: # xử lí các item trong Doc
            item = ms.Item(i)
            if item_type in item.ObjectName:
                item.Visible = show # ẨN Item
            if 'AcDbBlockReference' in item.ObjectName: # lấy ds tên block referent
                block_names.append(item.Name)
        except Exception as ex:
            pass
    # xử lí các BLOCK
    blocks = doc.Blocks
    for block in blocks:
        try: 
            if block.Name in block_names:
                show_hide_items_in_block(block,item_type = item_type, show = show)
        except Exception as ex:
            # print(ex)
            pass 
    doc.Regen (True)

def show_hide_items_in_block(block,item_type = "Text",show =False): 
    for i in range(block.Count):
        try: # xử lí các item trong block
            item = block.Item(i)
            if item_type in item.ObjectName:
                item.Visible = show # ẨN item
            if 'AcDbBlockReference' in item.ObjectName: # xử lí tiếp các Block con
                show_hide_items_in_block(item,item_type = item_type, show = show)
        except Exception as ex:
            # print (ex)
            pass


def create_list_dic_block(doc):
    """Tạo danh sách Dictionary các block"""
    global list_dic_block,list_block_name
    blocks = doc.Blocks
    dic = {}
    for i in range(blocks.Count):
        block = blocks[i]
        try: 
            dic[i] = [block,False]   
            list_dic_block.append(dic)
            list_block_name.append(block.Name)
        except Exception as ex:
            pass 
def create_list_dic_block_2(doc):
    """Tạo danh sách Dictionary các block"""    
    list_dic_block = []
    list_block_name = []
    blocks = doc.Blocks
    dic = {}
    for i in range(blocks.Count):
        block = blocks[i]
        try: 
            dic[i] = [block,False]   
            list_dic_block.append(dic)
            list_block_name.append(block.Name)
        except Exception as ex:
            pass 
    return list_dic_block,list_block_name

def set_color_items_in_doc(doc,item_type = "Text",in_layer = None,R = 192,G = 192 , B = 192):
    """Thiết lập màu đối tượng trong Document"""
    global list_dic_block,list_block_name,count
    count = 0
    ms = doc.ModelSpace
    block_names = []
    color = ms.Item(0).TrueColor
    color.SetRGB(R,G,B)
    item = None
    layer = None
    for i in range(ms.Count):
        try: # xử lí các item trong Doc
            item = ms.Item(i)
            layer = item.layer
            if item_type.lower() in item.ObjectName.lower():
                if in_layer != None:
                    if in_layer.lower() in layer.lower():
                        item.TrueColor = color
                        count += 1
                else:
                    item.TrueColor = color      
                    count += 1              
                if 'Dimension'.lower() in item.ObjectName.lower():
                    if in_layer != None:
                        if in_layer.lower() in layer.lower():
                            try:
                                item.DimensionLineColor = color.ColorIndex
                            except:
                                pass
                            try:
                                item.ExtensionLineColor = color.ColorIndex
                            except:
                                pass
                            try:
                                item.TextColor = color.ColorIndex
                            except:
                                pass
                            count += 1
                    else:
                        try:
                            item.DimensionLineColor = color.ColorIndex
                        except:
                            pass
                        try:
                            item.ExtensionLineColor = color.ColorIndex
                        except:
                            pass
                        try:
                            item.TextColor = color.ColorIndex
                        except:                        
                            pass
                        count += 1
            if 'BlockReference'.lower() in item.ObjectName.lower(): # lấy ds tên block referent
                # block_names.append(item.Name)
                try:
                    if item.Name in list_block_name:
                        print(item.Name)
                        ii = block_names.index(item.Name)
                        block = list_dic_block[ii][ii][0]
                        flag = list_dic_block[ii][ii][1]
                        print(flag)
                        if flag ==  False :
                            set_color_items_in_block(block,item_type = item_type, in_layer = in_layer, R = R,G = G , B = B)
                            list_dic_block[ii][ii][1] = True
                except Exception as ex:
                    print(ex)
                    pass
        except Exception as ex:
            print(ex)
            pass    
    print(f"Đã xử lí {count} đối tượng {item_type}")
    doc.Regen (True)
def set_color_items_in_block(block,item_type = "Text",in_layer = None, R = 192,G = 192 , B = 192): 
    """Thiết lập màu đối tượng trong Block"""
    global list_dic_block,list_block_name, count
    color = block.Item(0).TrueColor
    color.SetRGB(R,G,B)
    item = None
    layer = None
    block_names = []
    for i in range(block.Count):
        try: # xử lí các item trong block
            item = block.Item(i)
            layer = item.layer
            if item_type.lower() in item.ObjectName.lower():
                if in_layer != None:
                    if in_layer.lower() in layer.lower(): 
                        item.TrueColor = color
                        count += 1
                else:
                    item.TrueColor = color
                    count += 1
                if 'Dimension'.lower() in item.ObjectName.lower():
                    if in_layer != None:
                        if in_layer.lower() in layer.lower():
                            try:
                                item.DimensionLineColor = color.ColorIndex
                            except:
                                pass
                            try:
                                item.ExtensionLineColor = color.ColorIndex
                            except:
                                pass
                            try:
                                item.TextColor = color.ColorIndex
                            except:
                                pass
                            count += 1
                    else:
                        try:
                            item.DimensionLineColor = color.ColorIndex
                        except:
                            pass
                        try:
                            item.ExtensionLineColor = color.ColorIndex
                        except:
                            pass
                        try:
                            item.TextColor = color.ColorIndex
                        except:                        
                            pass
                        count += 1
            elif 'BlockReference'.lower() in item.ObjectName.lower(): # xử lí tiếp các Block con
                try:
                    if item.Name in list_block_name:
                        ii = block_names.index(item.Name)
                        block = list_dic_block[ii][ii][0]
                        flag = list_dic_block[ii][ii][1]
                        if flag ==  False :
                            set_color_items_in_block(block,item_type = item_type, in_layer = in_layer, R = R,G = G , B = B)
                            list_dic_block[ii][ii][1] = True
                except:
                    pass
        except Exception as ex:
            # print (ex)
            pass
# GLOBAL VAR #------------------------------------------------------------------#
list_dic_block = []
list_block_name = []
color = None
count = None

#-------------------------------------------------------------------#
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

    # LAYOUT
    # layouts = list(cad_doc.Layouts)
    # [print(f"{i} : {layouts[i].Name}") for i in range(len(layouts))]    
    # logging.info([layout.Name for layout in layouts])
    # # HỎI CHỌN LAYOUT
    # while True:
    #     try:
    #         ask_layout = input("Vui lòng chọn Layout: ")    
    #         if ask_layout == '0' or int(ask_layout):
    #             layout = layouts[int(ask_layout)]
    #             break
    #     except:
    #         continue
    # logging.info(layout.Name)

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
    
    # print("Bạn cần active vào Viewport bạn muốn xử lí")
    # active_viewport = cad_doc.ActivePViewport
    # logging.info(f"{active_viewport.ObjectName} {round(active_viewport.Width)} x {round(active_viewport.Height)}")
    
    # #  VIEWPORT LAYER
    # vp_doc = active_viewport.Document
    # logging.info(f"{vp_doc.Name}")

    # layers = list(active_viewport.Layer)
    # [print(layer.Name) for layer in layers]


    # GLOBAL LIST DIC BLOCK
    create_list_dic_block(cad_doc)

    list_process = ['Set Color','Show hide']
    print('\n'.join([f"{i} : {list_process[i]}" for i in range(len(list_process))]))
    while True:
        try:
            ask_process = list_process[int(input('\tVui lòng chọn Tiến trình: '))]
            if ask_process:
                break
        except:
            pass
    
    
    while True:
        # CHỌN LOẠI
        list_types = ['Text','Dimension','Leader','Hatch','Line','Polyline','Arc','Circle','Rectange','Elipse','Wipeout','MText']
        print('\n'.join([f"{i} : {list_types[i]}" for i in range(len(list_types))]))
        while True:
            try:
                ask_type = [list_types[int(i)] for i in input('\tVui lòng chọn Loại đối tượng: ').split(',')]
                if ask_type:
                    break
            except:
                pass
        # XỬ LÍ ẨN HIỆN
        if ask_process == 'Show hide':
            print('\n'.join(["0 : Show","1 : Hide"]))
            while True:
                try:
                    ask_show_hide = input('\tBạn muốn show or hide ? ')
                    if ask_show_hide == "0":
                        show = True
                        break
                    elif  ask_show_hide == "1":
                        show = False
                        break
                except:
                    pass
            time_report = [time.strftime("%d-%m-%y %H:%M:%S")]
            show_hide_items_in_doc(cad_doc,item_type = ask_type, show = show)  
            time_report.append(time.strftime("%d-%m-%y %H:%M:%S"))
            print (time_report)
        # XỬ LÍ MÀU
        if ask_process == 'Set Color':
            while True:
                try:
                    ask_R = int(input("\tRED index for color: " ))
                    ask_G = int(input("\tGREEN index for color: " ))
                    ask_B = int(input("\tBLUE index for color: " ))
                    if ask_R <= 255 and ask_R >=0 and ask_G <= 255 and ask_G >=0 and ask_B <= 255 and ask_B >=0:
                        break
                except:
                    pass        
            
            while True:
                try:
                    ask_in_layer = input('\tVui lòng nhập từ khóa Layer (tùy chọn): ')
                    if ask_in_layer == '':
                        ask_in_layer = None
                        break
                    else:
                        break
                except:
                    ask_in_layer = None
                    break            
            
            time_report = [time.strftime("%d-%m-%y %H:%M:%S")]            
            if len(ask_type) > 1:
                for type in ask_type:
                    set_color_items_in_doc(cad_doc,item_type = type, in_layer = ask_in_layer, R = ask_R, G = ask_G, B = ask_B)  
            
            elif len(ask_type) == 1:
                set_color_items_in_doc(cad_doc,item_type = ask_type[0], in_layer = ask_in_layer, R = ask_R, G = ask_G, B = ask_B)  
            time_report.append(time.strftime("%d-%m-%y %H:%M:%S"))
            
            print (time_report)
        
        print('\n'.join(["C : Tiếp tục","Q : Thoát"]))
        while True:
            try:
                ask_quit = input('Bạn muốn dừng trình xử lí? ').lower()
                if ask_quit == 'q'or ask_quit == 'c':
                    break                
            except:
                pass
        if ask_quit == 'q':
            break
