#!/usr/bin/env python3
from __future__ import annotations
import os, subprocess, sys, tempfile, time
from pathlib import Path

DISPLAY=os.environ.get('DISPLAY', ':99')
CAP_W=int(os.environ.get('CAPTURE_WIDTH','1365'))
CAP_H=int(os.environ.get('CAPTURE_HEIGHT','900'))
FPS=int(os.environ.get('CAPTURE_FPS','30'))
SECONDS=int(os.environ.get('RECORD_SECONDS','12'))

def quiet(cmd):
    return subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)

def capture(cmd):
    return subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, check=False)

def wait_window(timeout=20):
    end=time.time()+timeout
    while time.time()<end:
        r=capture(['xdotool','search','--onlyvisible','--class','libreoffice'])
        if r.returncode==0 and r.stdout.strip(): return r.stdout.splitlines()[0]
        r=capture(['xdotool','search','--onlyvisible','--name','.*'])
        if r.returncode==0:
            for wid in r.stdout.splitlines():
                n=capture(['xdotool','getwindowname',wid])
                if n.returncode==0 and 'calc' in n.stdout.lower(): return wid
        time.sleep(.5)
    raise RuntimeError('LibreOffice Calc window was not detected on DISPLAY '+DISPLAY)

def key(k):
    r=capture(['xdotool','key','--clearmodifiers',k])
    if r.returncode: raise RuntimeError('xdotool key failed: '+r.stdout)

def type_text(s,delay=30):
    r=capture(['xdotool','type','--clearmodifiers','--delay',str(delay),s])
    if r.returncode: raise RuntimeError('xdotool type failed: '+r.stdout)

def main():
    if len(sys.argv)!=3: raise SystemExit('Usage: python record_calc.py workbook.xlsx output.mp4')
    xlsx=Path(sys.argv[1]).resolve(); video=Path(sys.argv[2]).resolve()
    if not xlsx.exists(): raise FileNotFoundError(xlsx)
    video.parent.mkdir(parents=True,exist_ok=True); video.unlink(missing_ok=True)
    formula=os.environ.get('EXCEL_FORMULA','').strip()
    formula_col=int(os.environ.get('FORMULA_COL','5'))
    demo_value=os.environ.get('DEMO_VALUE','')

    quiet(['pkill','-f','soffice.bin']); quiet(['pkill','-f','soffice']); time.sleep(1)

    ff=['ffmpeg','-hide_banner','-loglevel','warning','-y','-f','x11grab','-video_size',f'{CAP_W}x{CAP_H}','-framerate',str(FPS),'-draw_mouse','1','-i',f'{DISPLAY}.0','-t',str(SECONDS),'-c:v','libx264','-preset','veryfast','-crf','23','-pix_fmt','yuv420p',str(video)]
    rec=subprocess.Popen(ff,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
    calc=None
    try:
        time.sleep(1)
        prof=Path(tempfile.mkdtemp(prefix='lo-profile-'))
        env=os.environ.copy(); env['DISPLAY']=DISPLAY; env['SAL_USE_VCLPLUGIN']='gen'
        calc=subprocess.Popen(['libreoffice','--nologo','--nodefault','--norestore','--nofirststartwizard',f'-env:UserInstallation=file://{prof}',str(xlsx)],env=env,stdout=subprocess.DEVNULL,stderr=subprocess.PIPE,text=True)
        wid=wait_window()
        quiet(['xdotool','windowactivate','--sync',wid]); quiet(['xdotool','windowraise',wid]); quiet(['xdotool','key','alt+F10']); time.sleep(1)
        key('ctrl+home')
        for _ in range(4): key('down')
        if demo_value:
            key('f2'); key('ctrl+a'); type_text(demo_value,45); key('Return'); time.sleep(.7)
        key('ctrl+home')
        for _ in range(4): key('down')
        for _ in range(formula_col-1): key('right')
        if formula:
            key('f2'); key('ctrl+a'); type_text(formula,25); key('Return'); time.sleep(2)
        time.sleep(max(1,SECONDS-7))
    finally:
        quiet(['xdotool','key','alt+F4']); time.sleep(1); quiet(['pkill','-f','soffice.bin']); quiet(['pkill','-f','soffice'])
        try: rec.wait(timeout=12)
        except subprocess.TimeoutExpired:
            rec.terminate()
            try: rec.wait(timeout=5)
            except subprocess.TimeoutExpired: rec.kill(); rec.wait()
    if not video.exists() or video.stat().st_size<10000:
        out=rec.stdout.read() if rec.stdout else ''
        raise RuntimeError(f'Calc recording was not created correctly: {video}\nFFmpeg output:\n{out}')
    probe=capture(['ffprobe','-v','error','-show_entries','format=duration,size','-show_entries','stream=codec_name,width,height','-of','default=noprint_wrappers=1',str(video)])
    if probe.returncode: raise RuntimeError('Recorded MP4 failed ffprobe:\n'+probe.stdout)
    print('Calc recording validated:\n'+probe.stdout)

if __name__=='__main__': main()
