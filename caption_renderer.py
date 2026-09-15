from pathlib import Path
from PIL import Image,ImageDraw,ImageFont
W,H=1080,1920
FONT=Path(r"C:\Windows\Fonts\arialbd.ttf")
if not FONT.exists(): FONT=Path(r"C:\Windows\Fonts\segoeuib.ttf")
def make_card(text,out):
    im=Image.new("RGBA",(W,H),(0,0,0,0)); d=ImageDraw.Draw(im)
    f=ImageFont.truetype(str(FONT),62)
    lines=text.split("\n"); y=1410
    for line in lines:
        bb=d.textbbox((0,0),line,font=f,stroke_width=4); x=(W-(bb[2]-bb[0]))//2
        d.text((x,y),line,font=f,fill=(255,255,255,255),stroke_width=5,stroke_fill=(0,0,0,255)); y+=72
    small=ImageFont.truetype(str(FONT),27)
    d.rounded_rectangle((105,1690,975,1750),22,fill=(0,0,0,175))
    d.text((540,1720),"LEARN VERSE  •  @LearnVerse9556",font=small,fill=(255,255,255,255),anchor="mm")
    im.save(out)
