from __future__ import annotations

import time
import pyautogui
from win32gui import SetForegroundWindow, GetWindowRect


def focus_excel(excel):
    SetForegroundWindow(excel.Hwnd)
    time.sleep(0.6)
    return GetWindowRect(excel.Hwnd)


def cell_center(excel, sheet, address):
    cell = sheet.Range(address)
    x = int(excel.ActiveWindow.PointsToScreenPixelsX(cell.Left + cell.Width / 2))
    y = int(excel.ActiveWindow.PointsToScreenPixelsY(cell.Top + cell.Height / 2))
    return x, y


def zoom_to_cell(excel, sheet, address, zoom=140):
    sheet.Activate()
    focus_excel(excel)
    try:
        excel.WindowState = -4137
        excel.ActiveWindow.Zoom = zoom
    except Exception:
        pass
    try:
        sheet.Range(address).Select()
        time.sleep(0.25)
    except Exception:
        pass
    time.sleep(0.9)


def restore_full_excel_view(excel, sheet, zoom=80):
    """Return Excel to A1 at the requested zoom so the final shot starts at A:G."""
    sheet.Activate()
    focus_excel(excel)
    try:
        excel.WindowState = -4137
        excel.ActiveWindow.Zoom = zoom
    except Exception:
        pass
    try:
        sheet.Range("A1").Select()
        pyautogui.hotkey("ctrl", "home")
        time.sleep(0.4)
        excel.ActiveWindow.ScrollColumn = 1
        excel.ActiveWindow.ScrollRow = 1
        time.sleep(0.5)
        excel.ActiveWindow.ScrollColumn = 1
        excel.ActiveWindow.ScrollRow = 1
    except Exception:
        pass
    time.sleep(1.2)


def click_address(excel, sheet, address):
    x, y = cell_center(excel, sheet, address)
    pyautogui.moveTo(x, y, duration=0.35)
    pyautogui.click()
    time.sleep(0.45)


def type_text(excel, sheet, address, text, zoom=140):
    zoom_to_cell(excel, sheet, address, zoom)
    x, y = cell_center(excel, sheet, address)
    pyautogui.moveTo(x, y, duration=0.35)
    pyautogui.click()
    time.sleep(0.45)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(text), interval=0.085)
    pyautogui.press("enter")
    time.sleep(1.2)


def type_formula(excel, sheet, address, formula, zoom=140):
    zoom_to_cell(excel, sheet, address, zoom)
    x, y = cell_center(excel, sheet, address)
    pyautogui.moveTo(x, y, duration=0.35)
    pyautogui.click()
    time.sleep(0.45)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(formula), interval=0.05)
    pyautogui.press("enter")
    time.sleep(2.0)


def click_checkbox(excel, sheet, row):
    # The real Form Control is linked to C{row}; clicking its cell works reliably.
    click_address(excel, sheet, f"C{row}")
    time.sleep(0.6)
