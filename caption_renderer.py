from __future__ import annotations
import sys
from pathlib import Path
from PIL import Image,ImageDraw,ImageFont

W,H=1080,1920
FONT="/Windows/Fonts/Arialbd.ttf"
if not Path(FONT).exists(): FONT="/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

def make_card(text,out):
    im=Image.new("RGBA",(W,H),(0,0,0,0))
    d=ImageDraw.Draw(im)
    f=ImageFont.truetype(FONT,64)
    # Reference-style white text + yellow emphasis simulated with a clean high-contrast block.
    lines=text.split("\n")
    y=1430
    for line in lines:
        box=d.textbbox((0,0),line,font=f,stroke_width=3)
        x=(W-(box[2]-box[0]))//2
        d.text((x,y),line,font=f,fill=(255,255,255,255),stroke_width=5,stroke_fill=(20,20,20,255))
        y+=76
    d.rounded_rectangle((120,1660,960,1740),25,fill=(0,0,0,120))
    d.text((W//2,1680),"Learn Verse  •  @LearnVerse9556",font=ImageFont.truetype(FONT,28),fill="white",anchor="mm")
    im.save(out)

def main():
    if len(sys.argv)!=3: raise SystemExit("Usage: caption_renderer.py text output.png")
    make_card(sys.argv[1],sys.argv[2])
if __name__=="__main__": main()
