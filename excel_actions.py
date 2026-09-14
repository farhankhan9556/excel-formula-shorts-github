from __future__ import annotations
import time, math
import pyautogui
import win32com.client as win32
from win32gui import GetForegroundWindow, GetWindowRect, SetForegroundWindow

def focus_excel(excel):
    hwnd=excel.Hwnd
    SetForegroundWindow(hwnd)
    time.sleep(1)
    return GetWindowRect(hwnd)

def cell_center(excel, sheet, row, col):
    # Excel Range.Left/Top are points within worksheet. Use Window.PointsToScreenPixelsX/Y.
    cell=sheet.Cells(row,col)
    x_points=cell.Left + cell.Width/2
    y_points=cell.Top + cell.Height/2
    x=int(excel.ActiveWindow.PointsToScreenPixelsX(x_points))
    y=int(excel.ActiveWindow.PointsToScreenPixelsY(y_points))
    return x,y

def click_cell(excel, sheet, row, col):
    x,y=cell_center(excel,sheet,row,col)
    pyautogui.moveTo(x,y,duration=0.25)
    pyautogui.click()
    time.sleep(0.5)

def click_checkbox(excel, sheet, row):
    # Form controls are positioned over column B. Use the center of the cell.
    x,y=cell_center(excel,sheet,row,2)
    pyautogui.moveTo(x,y,duration=0.25)
    pyautogui.click()
    time.sleep(0.6)

def type_formula(excel, sheet, row, col, formula):
    click_cell(excel,sheet,row,col)
    pyautogui.write(formula,interval=0.03)
    pyautogui.press("enter")
    time.sleep(1)

def select_range(excel,sheet,r1,c1,r2,c2):
    click_cell(excel,sheet,r1,c1)
    x1,y1=cell_center(excel,sheet,r1,c1)
    x2,y2=cell_center(excel,sheet,r2,c2)
    pyautogui.moveTo(x1,y1,duration=.2)
    pyautogui.dragTo(x2,y2,duration=.5,button="left")
    time.sleep(.5)
