import json, os, sys
import numpy as np
import matplotlib
import pyautocad
from pyautocad import *
import comtypes.client
import time

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