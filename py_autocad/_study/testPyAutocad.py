import pyautocad
from pyautocad import Autocad, ACAD, APoint
from pyautocad.utils import timing

acad = Autocad()

def sayhello():
    print('Hello')
#acad.prompt("Hello, Autocad from Python\n")# phai mo Autocad va file len

# p1 = APoint(0, 0)
# p2 = APoint(50, 25)
# for i in range(5):
#     text = acad.model.AddText('Hi %s!' % i, p1, 2.5)
#     acad.model.AddLine(p1, p2)
#     acad.model.AddCircle(p1, 10)
#     p1.y += 10

# dp = APoint(10, 0)
# for text in acad.iter_objects('Text'):
#     print('text: %s at: %s' % (text.TextString, text.InsertionPoint))
#     text.InsertionPoint = APoint(text.InsertionPoint) + dp

# for obj in acad.iter_objects(['Circle', 'Line']):
#     print(obj.ObjectName)

#check layer
# print("Current layer:",str(something.Layer))
# print("Current layer:",str(something.Linetype))
# print("Current layer:",str(something.LinetypeScale))
# print("Current layer:",str(something.Lineweight))
# print("Current layer:",str(something.Thickness))