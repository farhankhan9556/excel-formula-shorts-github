#!/usr/bin/env python3
from __future__ import annotations
import subprocess
import sys
import tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

W, H, FPS, DUR = 1080, 1920, 30, 14
FONT_DIR = Path("/usr/share/fonts/truetype/dejavu")
REG = FONT_DIR / "DejaVuSans.ttf"
BOLD = FONT_DIR / "DejaVuSans-Bold.ttf"

def font(size, bold=False):
    path = BOLD if bold else REG
    return ImageFont.truetype(str(path), size) if path.exists() else ImageFont.load_default()

def wrap(draw, text, fnt, max_width):
    lines, current = [], ""
    for word in str(text).split():
        candidate = word if not current else current + " " + word
        if draw.textbbox((0, 0), candidate, font=fnt)[2] <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines

def hook(problem, title, path):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((40, 45, 1040, 430), 35, fill=(17, 24, 39, 248))
    d.text((82, 82), "EXCEL QUICK TIP", font=font(30, True), fill=(100, 185, 255, 255))
    y = 140
    for line in wrap(d, problem, font(43, True), 880)[:3]:
        d.text((82, y), line, font=font(43, True), fill="white")
        y += 58
    d.text((82, 350), "Watch the real spreadsheet solve it →", font=font(26), fill=(210, 220, 230, 255))
    d.rounded_rectangle((65, 455, 1015, 515), 18, fill=(35, 115, 200, 235))
    title_lines = wrap(d, title, font(25, True), 850)
    d.text((95, 471), title_lines[0] if title_lines else title, font=font(25, True), fill="white")
    im.save(path)

def formula_card(formula, result, path):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((40, 1370, 1040, 1710), 30, fill=(255, 255, 255, 248))
    d.text((82, 1405), "FORMULA", font=font(27, True), fill=(50, 90, 130, 255))
    y = 1460
    for line in wrap(d, formula, font(29, True), 880)[:2]:
        d.text((82, y), line, font=font(29, True), fill=(20, 30, 40, 255))
        y += 42
    d.rounded_rectangle((72, 1570, 1008, 1675), 22, fill=(35, 130, 80, 245))
    d.text((100, 1600), f"✓ Result: {result}", font=font(36, True), fill="white")
    im.save(path)

def cta(logo, path):
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((40, 70, 1040, 480), 36, fill=(17, 24, 39, 250))
    lp = Path(logo)
    if lp.exists():
        try:
            lg = Image.open(lp).convert("RGBA")
            lg.thumbnail((175, 175))
            im.alpha_composite(lg, (82, 165))
        except Exception as exc:
            print("Logo warning:", exc)
    d.text((300, 145), "FOLLOW FOR MORE", font=font(38, True), fill="white")
    d.text((300, 215), "Excel & Workplace Tips", font=font(29), fill=(215, 225, 235, 255))
    d.text((300, 285), "Learn Verse", font=font(50, True), fill="white")
    d.text((300, 360), "@LearnVerse9556", font=font(27), fill=(170, 195, 220, 255))
    im.save(path)

def main():
    if len(sys.argv) != 9:
        raise SystemExit("Usage: render_video.py raw.mp4 voice.mp3 out.mp4 problem title formula result logo.png")

    raw, voice, out = map(Path, sys.argv[1:4])
    problem, title, formula, result = sys.argv[4:8]
    logo = Path(sys.argv[8])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.unlink(missing_ok=True)

    for p, label in ((raw, "raw video"), (voice, "voice")):
        if not p.exists() or p.stat().st_size < 1000:
            raise RuntimeError(f"Missing or invalid {label}: {p}")

    with tempfile.TemporaryDirectory(prefix="lv-overlay-") as td:
        td = Path(td)
        hook_png = td / "hook.png"
        formula_png = td / "formula.png"
        cta_png = td / "cta.png"
        hook(problem, title, hook_png)
        formula_card(formula, result, formula_png)
        cta(logo, cta_png)

        filt = (
            "[0:v]scale=960:850:force_original_aspect_ratio=decrease,"
            "pad=960:850:(ow-iw)/2:(oh-ih)/2:white,setsar=1[calc];"
            "color=c=white:s=1080x1920:r=30:d=14[bg];"
            "[bg][calc]overlay=60:540[base];"
            "[1:v]format=rgba[h];"
            "[base][h]overlay=0:0:enable='between(t,0,4.5)'[v1];"
            "[2:v]format=rgba[f];"
            "[v1][f]overlay=0:0:enable='between(t,3,11)'[v2];"
            "[3:v]format=rgba[c];"
            "[v2][c]overlay=0:0:enable='between(t,10.5,14)'[vout];"
            "[4:a]aresample=async=1,apad,atrim=0:14,loudnorm=I=-16:TP=-1.5:LRA=11[a]"
        )

        cmd = [
            "ffmpeg", "-hide_banner", "-y",
            "-stream_loop", "-1", "-i", str(raw),
            "-loop", "1", "-i", str(hook_png),
            "-loop", "1", "-i", str(formula_png),
            "-loop", "1", "-i", str(cta_png),
            "-i", str(voice),
            "-filter_complex", filt,
            "-map", "[vout]",
            "-map", "[a]",
            "-t", str(DUR),
            "-r", str(FPS),
            "-c:v", "libx264",
            "-preset", "veryfast",
            "-crf", "22",
            "-pix_fmt", "yuv420p",
            "-c:a", "aac",
            "-b:a", "128k",
            "-ar", "48000",
            "-movflags", "+faststart",
            str(out),
        ]
        r = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        print(r.stdout)
        if r.returncode:
            raise RuntimeError(f"Final FFmpeg render failed ({r.returncode})")

    if not out.exists() or out.stat().st_size < 50000:
        raise RuntimeError(f"Final MP4 invalid: {out}")

    r = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-show_entries", "stream=codec_name,width,height,r_frame_rate",
            "-show_entries", "format=duration,size",
            "-of", "default=noprint_wrappers=1",
            str(out),
        ],
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
    )
    if r.returncode:
        raise RuntimeError(f"Final MP4 failed ffprobe:\n{r.stdout}")
    print("FINAL VIDEO VALIDATED")
    print(r.stdout)

if __name__ == "__main__":
    main()
