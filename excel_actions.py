from __future__ import annotations
import time
import pyautogui
from win32gui import SetForegroundWindow, GetWindowRect

def focus_excel(excel):
    SetForegroundWindow(excel.Hwnd)
    time.sleep(1)
    return GetWindowRect(excel.Hwnd)

def cell_center(excel,sheet,row,col):
    cell=sheet.Cells(row,col)
    x=int(excel.ActiveWindow.PointsToScreenPixelsX(cell.Left+cell.Width/2))
    y=int(excel.ActiveWindow.PointsToScreenPixelsY(cell.Top+cell.Height/2))
    return x,y

def click_cell(excel,sheet,row,col):
    x,y=cell_center(excel,sheet,row,col)
    pyautogui.moveTo(x,y,duration=.25); pyautogui.click(); time.sleep(.45)

def type_text(excel,sheet,row,col,text):
    click_cell(excel,sheet,row,col)
    pyautogui.write(str(text),interval=.05); pyautogui.press('enter'); time.sleep(.8)

def type_formula(excel,sheet,row,col,formula):
    click_cell(excel,sheet,row,col)
    pyautogui.write(formula,interval=.025); pyautogui.press('enter'); time.sleep(1.5)

def drag_select(excel,sheet,r1,c1,r2,c2):
    x1,y1=cell_center(excel,sheet,r1,c1); x2,y2=cell_center(excel,sheet,r2,c2)
    pyautogui.moveTo(x1,y1,duration=.2); pyautogui.dragTo(x2,y2,duration=.7,button='left'); time.sleep(.5)
