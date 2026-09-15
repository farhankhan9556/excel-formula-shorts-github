from __future__ import annotations
import subprocess,time
from pathlib import Path
from excel_actions import focus_excel

def record_excel(excel,out,duration=30):
    x1,y1,x2,y2=focus_excel(excel)
    # Remove a small outer border while retaining the full Excel ribbon/workbook.
    x=max(0,x1); y=max(0,y1); w=max(2,x2-x); h=max(2,y2-y)
    if w%2: w-=1
    if h%2: h-=1
    vf=f"crop={w}:{h}:{x}:{y},scale=1080:1920:force_original_aspect_ratio=decrease,pad=1080:1920:(ow-iw)/2:(oh-ih)/2,setsar=1"
    cmd=["ffmpeg","-hide_banner","-loglevel","warning","-y","-f","gdigrab","-framerate","30","-draw_mouse","1","-i","desktop","-vf",vf,"-t",str(duration),"-c:v","libx264","-preset","veryfast","-crf","19","-pix_fmt","yuv420p",str(out)]
    p=subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
    time.sleep(1)
    return p
