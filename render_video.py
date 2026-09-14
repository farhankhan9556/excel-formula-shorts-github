#!/usr/bin/env python3
from __future__ import annotations
import subprocess,sys,tempfile
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
W,H,FPS,DUR=1080,1920,30,14
FD=Path('/usr/share/fonts/truetype/dejavu'); REG=FD/'DejaVuSans.ttf'; BOLD=FD/'DejaVuSans-Bold.ttf'
def ft(n,b=False): return ImageFont.truetype(str(BOLD if b else REG),n) if (BOLD if b else REG).exists() else ImageFont.load_default()
def wrap(d,t,f,m):
    out=[]; cur=''
    for w in str(t).split():
        s=w if not cur else cur+' '+w
        if d.textbbox((0,0),s,font=f)[2]<=m: cur=s
        else:
            if cur: out.append(cur)
            cur=w
    if cur: out.append(cur)
    return out
def hook(problem,title,p):
    im=Image.new('RGBA',(W,H),(0,0,0,0)); d=ImageDraw.Draw(im); d.rounded_rectangle((40,45,1040,430),35,fill=(17,24,39,248)); d.text((82,82),'EXCEL QUICK TIP',font=ft(30,1),fill=(100,185,255,255)); y=140
    for line in wrap(d,problem,ft(43,1),880)[:3]: d.text((82,y),line,font=ft(43,1),fill='white'); y+=58
    d.text((82,350),'Watch the real spreadsheet solve it →',font=ft(26),fill=(210,220,230,255)); d.rounded_rectangle((65,455,1015,515),18,fill=(35,115,200,235)); d.text((95,471),wrap(d,title,ft(25,1),850)[0],font=ft(25,1),fill='white'); im.save(p)
def formula_card(formula,result,p):
    im=Image.new('RGBA',(W,H),(0,0,0,0)); d=ImageDraw.Draw(im); d.rounded_rectangle((40,1370,1040,1710),30,fill=(255,255,255,248)); d.text((82,1405),'FORMULA',font=ft(27,1),fill=(50,90,130,255)); y=1460
    for line in wrap(d,formula,ft(29,1),880)[:2]: d.text((82,y),line,font=ft(29,1),fill=(20,30,40,255)); y+=42
    d.rounded_rectangle((72,1570,1008,1675),22,fill=(35,130,80,245)); d.text((100,1600),f'✓ Result: {result}',font=ft(36,1),fill='white'); im.save(p)
def cta(logo,p):
    im=Image.new('RGBA',(W,H),(0,0,0,0)); d=ImageDraw.Draw(im); d.rounded_rectangle((40,70,1040,480),36,fill=(17,24,39,250)); lp=Path(logo)
    if lp.exists():
        try: lg=Image.open(lp).convert('RGBA'); lg.thumbnail((175,175)); im.alpha_composite(lg,(82,165))
        except Exception as e: print('Logo warning:',e)
    d.text((300,145),'FOLLOW FOR MORE',font=ft(38,1),fill='white'); d.text((300,215),'Excel & Workplace Tips',font=ft(29),fill=(215,225,235,255)); d.text((300,285),'Learn Verse',font=ft(50,1),fill='white'); d.text((300,360),'@LearnVerse9556',font=ft(27),fill=(170,195,220,255)); im.save(p)
def main():
    if len(sys.argv)!=9: raise SystemExit('Usage: render_video.py raw.mp4 voice.mp3 out.mp4 problem title formula result logo.png')
    raw,voice,out=map(Path,sys.argv[1:4]); problem,title,formula,result=sys.argv[4:8]; logo=Path(sys.argv[8]); out.parent.mkdir(parents=True,exist_ok=True); out.unlink(missing_ok=True)
    for p,label in ((raw,'raw video'),(voice,'voice')):
        if not p.exists() or p.stat().st_size<1000: raise RuntimeError(f'Missing/invalid {label}: {p}')
    with tempfile.TemporaryDirectory(prefix='lv-overlay-') as td:
        td=Path(td); h=td/'hook.png'; f=td/'formula.png'; c=td/'cta.png'; hook(problem,title,h); formula_card(formula,result,f); cta(logo,c)
        filt="[0:v]scale=960:850:force_original_aspect_ratio=decrease,pad=960:850:(ow-iw)/2:(oh-ih)/2:white,setsar=1[calc];color=c=white:s=1080x1920:r=30:d=14[bg];[bg][calc]overlay=60:540[base];[1:v]format=rgba[h];[base][h]overlay=0:0:enable='between(t,0,4.5)'[v1];[2:v]format=rgba[f];[v1][f]overlay=0:0:enable='between(t,3,11)'[v2];[3:v]format=rgba[c];[v2][c]overlay=0:0:enable='between(t,10.5,14)'[vout];[4:a]aresample=async=1,loudnorm=I=-16:TP=-1.5:LRA=11[a]"
        cmd=['ffmpeg','-hide_banner','-y','-stream_loop','-1','-i',str(raw),'-loop','1','-i',str(h),'-loop','1','-i',str(f),'-loop','1','-i',str(c),'-i',str(voice),'-filter_complex',filt,'-map','[vout]','-map','[a]','-t',str(DUR),'-r',str(FPS),'-c:v','libx264','-preset','veryfast','-crf','22','-pix_fmt','yuv420p','-c:a','aac','-b:a','128k','-ar','48000','-movflags','+faststart',str(out)]
        r=subprocess.run(cmd,text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT); print(r.stdout)
        if r.returncode: raise RuntimeError(f'Final FFmpeg render failed ({r.returncode})')
    if not out.exists() or out.stat().st_size<50000: raise RuntimeError(f'Final MP4 invalid: {out}')
    r=subprocess.run(['ffprobe','-v','error','-show_entries','stream=codec_name,width,height,r_frame_rate','-show_entries','format=duration,size','-of','default=noprint_wrappers=1',str(out)],text=True,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
    if r.returncode: raise RuntimeError('Final MP4 failed ffprobe:\n'+r.stdout)
    print('FINAL VIDEO VALIDATED\n'+r.stdout)
if __name__=='__main__': main()
