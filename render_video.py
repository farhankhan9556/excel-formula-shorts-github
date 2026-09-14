#!/usr/bin/env python3
from __future__ import annotations
import subprocess, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

W,H=1080,1920
FPS=30
DURATION=14
FONT_DIR=Path("/usr/share/fonts/truetype/dejavu")
REG=FONT_DIR/"DejaVuSans.ttf"
BOLD=FONT_DIR/"DejaVuSans-Bold.ttf"

def font(size,bold=False):
    p=BOLD if bold else REG
    return ImageFont.truetype(str(p),size) if p.exists() else ImageFont.load_default()

def wrap(draw,text,f,max_width):
    words=str(text).split()
    lines=[]; cur=""
    for word in words:
        test=word if not cur else cur+" "+word
        if draw.textbbox((0,0),test,font=f)[2] <= max_width:
            cur=test
        else:
            if cur: lines.append(cur)
            cur=word
    if cur: lines.append(cur)
    return lines

def base_overlay():
    return Image.new("RGBA",(W,H),(0,0,0,0))

def make_hook(problem,title,path):
    im=base_overlay(); d=ImageDraw.Draw(im)
    d.rounded_rectangle((40,45,1040,485),35,fill=(15,23,42,248))
    d.text((85,80),"EXCEL QUICK TIP",font=font(30,True),fill=(90,190,255,255))
    f=font(46,True); y=145
    for line in wrap(d,problem,f,870)[:3]:
        d.text((85,y),line,font=f,fill="white"); y+=62
    d.text((85,370),"Watch the real spreadsheet solve it →",font=font(27),fill=(215,225,235,255))
    d.rounded_rectangle((70,505,1010,565),22,fill=(35,115,200,240))
    tf=font(25,True)
    line=wrap(d,title,tf,850)[0]
    d.text((100,520),line,font=tf,fill="white")
    im.save(path)

def make_formula(formula,result,path):
    im=base_overlay(); d=ImageDraw.Draw(im)
    d.rounded_rectangle((40,1370,1040,1710),30,fill=(255,255,255,250))
    d.text((80,1410),"FORMULA",font=font(28,True),fill=(40,85,130,255))
    ff=font(30,True); y=1460
    for line in wrap(d,formula,ff,880)[:2]:
        d.text((80,y),line,font=ff,fill=(20,30,40,255)); y+=43
    d.rounded_rectangle((70,1560,1010,1670),24,fill=(35,125,80,245))
    d.text((100,1590),f"✓ Result: {result}",font=font(38,True),fill="white")
    im.save(path)

def make_cta(logo_path,path):
    im=base_overlay(); d=ImageDraw.Draw(im)
    d.rounded_rectangle((40,80,1040,485),35,fill=(15,23,42,250))
    lp=Path(logo_path)
    if lp.exists():
        try:
            logo=Image.open(lp).convert("RGBA")
            logo.thumbnail((175,175))
            im.alpha_composite(logo,(85,165))
        except Exception as exc:
            print(f"WARNING: logo could not be loaded: {exc}")
    d.text((300,145),"FOLLOW FOR MORE",font=font(38,True),fill="white")
    d.text((300,210),"Excel & Workplace Tips",font=font(29),fill=(215,225,235,255))
    d.text((300,275),"Learn Verse",font=font(48,True),fill="white")
    d.text((300,345),"@LearnVerse9556",font=font(27),fill=(175,195,220,255))
    im.save(path)

def run_ffmpeg(raw,voice,out,hook,formula,cta):
    for p,label,minsize in [(raw,"raw Calc video",10000),(voice,"voice",1000),(hook,"hook overlay",100),(formula,"formula overlay",100),(cta,"CTA overlay",100)]:
        if not Path(p).exists() or Path(p).stat().st_size<minsize:
            raise RuntimeError(f"Invalid {label}: {p}")
    out=Path(out); out.parent.mkdir(parents=True,exist_ok=True)
    # Inputs: 0 raw video, 1 hook PNG, 2 formula PNG, 3 CTA PNG, 4 audio.
    graph=(
        "[0:v]scale=960:850:force_original_aspect_ratio=decrease,"
        "pad=960:850:(ow-iw)/2:(oh-ih)/2:color=white,setsar=1[calc];"
        "[calc]null[calc2];"
        "[1:v]format=rgba[hook];"
        "[2:v]format=rgba[formula];"
        "[3:v]format=rgba[cta];"
        "[calc2][hook]overlay=60:520:enable='between(t,0,4.5)'[v1];"
        "[v1][formula]overlay=0:0:enable='between(t,3,11)'[v2];"
        "[v2][cta]overlay=0:0:enable='between(t,10.5,14)'[vout];"
        "[4:a]aresample=async=1,loudnorm=I=-16:TP=-1.5:LRA=11[aout]"
    )
    cmd=[
        "ffmpeg","-hide_banner","-y",
        "-stream_loop","-1","-i",str(raw),
        "-loop","1","-framerate",str(FPS),"-i",str(hook),
        "-loop","1","-framerate",str(FPS),"-i",str(formula),
        "-loop","1","-framerate",str(FPS),"-i",str(cta),
        "-i",str(voice),
        "-filter_complex",graph,
        "-map","[vout]","-map","[aout]",
        "-t",str(DURATION),
        "-r",str(FPS),
        "-c:v","libx264","-preset","veryfast","-crf","23","-pix_fmt","yuv420p",
        "-c:a","aac","-b:a","128k","-ar","48000",
        "-movflags","+faststart",
        str(out)
    ]
    print("Running FFmpeg:")
    print(" ".join(map(str,cmd)))
    result=subprocess.run(cmd,text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
    print(result.stdout)
    if result.returncode!=0:
        raise RuntimeError(f"FFmpeg failed with exit code {result.returncode}")
    if not out.exists() or out.stat().st_size<50000:
        raise RuntimeError("FFmpeg did not create a valid final MP4.")
    subprocess.run(["ffprobe","-v","error","-show_entries","format=duration,size","-of","default=noprint_wrappers=1",str(out)],check=True)
    print(f"VIDEO SUCCESS: {out} ({out.stat().st_size} bytes)")

def main():
    if len(sys.argv)!=9:
        raise SystemExit('Usage: python render_video.py raw.mp4 voice.mp3 output.mp4 "problem" "title" "formula" "result" logo.png')
    raw,voice,out,problem,title,formula,result,logo=sys.argv[1:]
    with tempfile.TemporaryDirectory(prefix="learnverse_") as td:
        td=Path(td)
        hook=td/"hook.png"; formula_png=td/"formula.png"; cta=td/"cta.png"
        make_hook(problem,title,hook)
        make_formula(formula,result,formula_png)
        make_cta(logo,cta)
        run_ffmpeg(raw,voice,out,hook,formula_png,cta)

if __name__=="__main__":
    main()
