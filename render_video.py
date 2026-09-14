#!/usr/bin/env python3
import subprocess
import sys
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

W, H = 1080, 1920
FPS = 30
DURATION = 14

def fnt(size, bold=False):
    p = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    return ImageFont.truetype(p, size) if Path(p).exists() else ImageFont.load_default()

def make_overlay(problem, formula, result, logo_path):
    tmp = Path("/tmp/excel_overlay")
    if tmp.exists():
        for x in tmp.glob("*.png"):
            x.unlink()
    tmp.mkdir(exist_ok=True)
    logo = Image.open(logo_path).convert("RGBA") if Path(logo_path).exists() else None

    for i in range(DURATION * FPS):
        t = i / FPS
        im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)

        # Problem hook
        if t < 2.5:
            d.rounded_rectangle((45, 55, 1035, 500), radius=35, fill=(17,24,39,248))
            d.text((90, 100), "THE PROBLEM", font=fnt(28, True), fill=(90,175,255,255))
            # Wrap problem over two lines.
            words = problem.split()
            lines, line = [], ""
            for word in words:
                test = (line + " " + word).strip()
                if d.textbbox((0,0), test, font=fnt(48, True))[2] > 860:
                    lines.append(line); line = word
                else:
                    line = test
            if line: lines.append(line)
            y = 165
            for ln in lines[:3]:
                d.text((90, y), ln, font=fnt(48, True), fill="white")
                y += 68
            d.text((90, 390), "Watch Excel solve it →", font=fnt(28), fill=(210,220,230,255))

        # Formula card
        if 2.0 <= t < 10.8:
            d.rounded_rectangle((45, 1390, 1035, 1710), radius=28, fill=(255,255,255,245))
            d.text((85, 1430), "FORMULA", font=fnt(24, True), fill=(55,85,125,255))
            d.text((85, 1480), formula, font=fnt(34, True), fill=(18,28,38,255))
            d.text((85, 1555), f"Outcome: {result}", font=fnt(42, True), fill=(30,120,75,255))
            d.text((85, 1630), "Real Excel-compatible workbook", font=fnt(24), fill=(95,105,115,255))

        # Result callout
        if 7.8 <= t < 10.8:
            d.rounded_rectangle((60, 1715, 1020, 1840), radius=25, fill=(30,120,75,245))
            d.text((100, 1745), f"✓ Result: {result}", font=fnt(40, True), fill="white")

        # Branding
        if t >= 10.5:
            d.rounded_rectangle((45, 95, 1035, 475), radius=35, fill=(17,24,39,250))
            if logo:
                lg = logo.copy()
                lg.thumbnail((170,170))
                im.alpha_composite(lg, (90, 165))
            d.text((295, 155), "FOLLOW FOR MORE", font=fnt(36, True), fill="white")
            d.text((295, 215), "Excel & Workplace Tips", font=fnt(29), fill=(215,225,235,255))
            d.text((295, 275), "Learn Verse", font=fnt(48, True), fill="white")
            d.text((295, 345), "Learn Anything. Anytime.", font=fnt(25), fill=(175,190,205,255))

        im.save(tmp / f"frame_{i:05d}.png")
    return tmp

def main():
    raw, voice, out, problem, title, formula, result, logo = sys.argv[1:]
    frames = make_overlay(problem, formula, result, logo)
    out = Path(out)

    # Keep the real Calc recording visible for the full video; overlay has transparency.
    cmd = [
        "ffmpeg","-y",
        "-stream_loop","-1","-i",raw,
        "-framerate",str(FPS),"-i",str(frames/"frame_%05d.png"),
        "-i",voice,
        "-filter_complex",
        "[0:v]scale=1080:1920:force_original_aspect_ratio=decrease,"
        "pad=1080:1920:(ow-iw)/2:(oh-ih)/2:color=white,setsar=1[base];"
        "[1:v]format=rgba[ov];"
        "[base][ov]overlay=0:0:format=auto[v];"
        "[2:a]loudnorm=I=-16:TP=-1.5:LRA=11[a]",
        "-map","[v]","-map","[a]","-t",str(DURATION),
        "-c:v","libx264","-preset","veryfast","-crf","23",
        "-pix_fmt","yuv420p","-c:a","aac","-b:a","128k",
        "-movflags","+faststart",str(out)
    ]
    subprocess.run(cmd, check=True)

if __name__ == "__main__":
    main()
