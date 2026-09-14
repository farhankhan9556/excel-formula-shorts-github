#!/usr/bin/env python3
from __future__ import annotations
import os, subprocess, sys, time
from pathlib import Path

if len(sys.argv)!=3:
    raise SystemExit("Usage: python record_calc.py workbook.xlsx raw_video.mp4")

xlsx=Path(sys.argv[1]).resolve()
video=Path(sys.argv[2]).resolve()
display=os.environ.get("DISPLAY",":99")
seconds=int(os.environ.get("RECORD_SECONDS","11"))
formula=os.environ.get("EXCEL_FORMULA","")
formula_col=int(os.environ.get("FORMULA_COL","5"))

if not xlsx.exists():
    raise FileNotFoundError(xlsx)
video.parent.mkdir(parents=True,exist_ok=True)

def xdotool(*args, check=False):
    return subprocess.run(["xdotool",*args],check=check,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)

subprocess.run(["pkill","-f","soffice"],check=False,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
time.sleep(1)

proc=subprocess.Popen(
    ["libreoffice","--nologo","--nodefault","--norestore","--nofirststartwizard",str(xlsx)],
    env={**os.environ,"DISPLAY":display},
    stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL
)

try:
    window=""
    for _ in range(20):
        result=subprocess.run(
            ["xdotool","search","--onlyvisible","--class","libreoffice"],
            env={**os.environ,"DISPLAY":display},
            text=True,capture_output=True
        )
        ids=result.stdout.strip().split()
        if ids:
            window=ids[0]
            break
        time.sleep(0.5)

    if not window:
        raise RuntimeError("LibreOffice Calc window was not found on Xvfb display.")

    xdotool("windowactivate","--sync",window)
    xdotool("key","alt+F10")
    time.sleep(1)

    # Go to A5.
    xdotool("key","ctrl+home")
    for _ in range(4):
        xdotool("key","down")
    time.sleep(0.5)

    # Demonstrate data entry in the first data row.
    # Select B5 and type the existing value back in with visible keystrokes.
    xdotool("key","right")
    xdotool("key","f2")
    xdotool("key","ctrl+a")
    xdotool("type","--delay","60","92")
    xdotool("key","Return")
    time.sleep(0.8)

    # Enter the genuine translated formula in the Formula cell F5/etc.
    xdotool("key","ctrl+home")
    for _ in range(4):
        xdotool("key","down")
    for _ in range(formula_col-1):
        xdotool("key","right")
    xdotool("key","f2")
    xdotool("key","ctrl+a")
    if formula:
        xdotool("type","--delay","20",formula)
        xdotool("key","Return")
    time.sleep(max(1,seconds-5))
finally:
    xdotool("key","alt+F4")
    time.sleep(1)
    subprocess.run(["pkill","-f","soffice"],check=False,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)

if not video.exists():
    raise RuntimeError("The Calc recording file was not created by the FFmpeg recorder.")
print(f"Calc automation completed: {video}")
