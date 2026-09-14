from __future__ import annotations
import subprocess,sys,time,os
import pyautogui
from pathlib import Path
import win32com.client as win32
from excel_actions import focus_excel,click_checkbox,click_cell,select_range,type_formula

def record_screen(out,seconds):
    out=Path(out)
    cmd=["ffmpeg","-hide_banner","-loglevel","warning","-y","-f","gdigrab","-framerate","30","-draw_mouse","1","-i","desktop","-t",str(seconds),"-c:v","libx264","-preset","veryfast","-crf","20","-pix_fmt","yuv420p",str(out)]
    return subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)

def run_checkbox(excel,book,video,out):
    sheet=book.Worksheets("Learn Verse")
    focus_excel(excel)
    rec=record_screen(out,30)
    time.sleep(1)
    # Match the reference rhythm: cursor -> first checkbox -> select group -> click examples.
    click_cell(excel,sheet,4,2)
    time.sleep(1.2)
    click_checkbox(excel,sheet,4)
    click_checkbox(excel,sheet,5)
    time.sleep(1)
    select_range(excel,sheet,4,2,12,2)
    time.sleep(1.2)
    for r in (4,5,7,10,12):
        click_checkbox(excel,sheet,r)
        # Status text is updated through COM to keep the demo deterministic.
        sheet.Cells(r,3).Value="Paid"
    for r in (6,8,9,11):
        sheet.Cells(r,3).Value="Unpaid"
    book.Save()
    time.sleep(15)
    rec.wait(timeout=15)
    return rec.returncode

def run_formula(excel,book,video,out):
    sheet=book.Worksheets("Learn Verse")
    focus_excel(excel)
    rec=record_screen(out,30)
    time.sleep(1)
    click_cell(excel,sheet,5,2)
    pyautogui.write("E103",interval=.08)
    pyautogui.press("enter")
    time.sleep(2)
    click_cell(excel,sheet,4,6)
    pyautogui.hotkey("ctrl","a")
    pyautogui.write(video["formula"],interval=.03)
    pyautogui.press("enter")
    time.sleep(4)
    rec.wait(timeout=15)
    return rec.returncode

if __name__=="__main__":
    raise SystemExit("Called by generate_windows.py")
