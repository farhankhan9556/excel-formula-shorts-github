#!/usr/bin/env python3

"""
Learn Verse - Excel Shorts Renderer

Creates a professional 1080x1920 YouTube Short from:
1. Real LibreOffice Calc screen recording
2. AI voice MP3
3. Problem/title/formula/result information
4. Learn Verse logo

This version avoids generating hundreds of PNG frames.
Only 3 transparent overlay images are created:
- Hook
- Formula/result
- CTA/branding

FFmpeg handles the timing and final encoding.
"""

from __future__ import annotations

import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


# ============================================================
# VIDEO SETTINGS
# ============================================================

WIDTH = 1080
HEIGHT = 1920
FPS = 30
DURATION = 14

CALC_X = 60
CALC_Y = 520
CALC_W = 960
CALC_H = 850


# ============================================================
# FONT HELPERS
# ============================================================

FONT_DIR = Path("/usr/share/fonts/truetype/dejavu")

FONT_REGULAR = FONT_DIR / "DejaVuSans.ttf"
FONT_BOLD = FONT_DIR / "DejaVuSans-Bold.ttf"


def get_font(size: int, bold: bool = False):
    path = FONT_BOLD if bold else FONT_REGULAR

    if path.exists():
        return ImageFont.truetype(str(path), size)

    return ImageFont.load_default()


# ============================================================
# TEXT WRAPPING
# ============================================================

def wrap_text(draw, text, font, max_width):
    words = str(text).split()

    lines = []
    current = ""

    for word in words:
        test = word if not current else current + " " + word

        bbox = draw.textbbox(
            (0, 0),
            test,
            font=font
        )

        text_width = bbox[2] - bbox[0]

        if text_width <= max_width:
            current = test
        else:
            if current:
                lines.append(current)

            current = word

    if current:
        lines.append(current)

    return lines


# ============================================================
# SAFE TEXT
# ============================================================

def safe_text(value):
    """
    Make text safe for Pillow.
    """
    if value is None:
        return ""

    return str(value).replace("\r", "").replace("\n", " ").strip()


# ============================================================
# CREATE HOOK OVERLAY
# ============================================================

def create_hook(problem, title, output):
    img = Image.new(
        "RGBA",
        (WIDTH, HEIGHT),
        (0, 0, 0, 0)
    )

    draw = ImageDraw.Draw(img)

    # Main card
    draw.rounded_rectangle(
        (40, 50, 1040, 430),
        radius=35,
        fill=(18, 25, 38, 245)
    )

    # Small label
    label_font = get_font(30, True)

    draw.text(
        (85, 90),
        "EXCEL QUICK TIP",
        font=label_font,
        fill=(120, 190, 255, 255)
    )

    # Problem
    problem_font = get_font(45, True)

    lines = wrap_text(
        draw,
        safe_text(problem),
        problem_font,
        880
    )

    y = 145

    for line in lines[:3]:
        draw.text(
            (85, y),
            line,
            font=problem_font,
            fill=(255, 255, 255, 255)
        )

        y += 62

    # Arrow/message
    small_font = get_font(27, False)

    draw.text(
        (85, 350),
        "Watch the real spreadsheet solve it →",
        font=small_font,
        fill=(205, 215, 230, 255)
    )

    # Title at bottom
    title_font = get_font(25, True)

    draw.rounded_rectangle(
        (70, 455, 1010, 510),
        radius=18,
        fill=(35, 115, 200, 235)
    )

    title_lines = wrap_text(
        draw,
        safe_text(title),
        title_font,
        850
    )

    title_text = title_lines[0] if title_lines else ""

    draw.text(
        (100, 468),
        title_text,
        font=title_font,
        fill=(255, 255, 255, 255)
    )

    img.save(output)


# ============================================================
# CREATE FORMULA OVERLAY
# ============================================================

def create_formula_overlay(formula, result, output):
    img = Image.new(
        "RGBA",
        (WIDTH, HEIGHT),
        (0, 0, 0, 0)
    )

    draw = ImageDraw.Draw(img)

    # Formula card
    draw.rounded_rectangle(
        (45, 1390, 1035, 1700),
        radius=30,
        fill=(255, 255, 255, 248)
    )

    # Header
    draw.text(
        (85, 1425),
        "FORMULA",
        font=get_font(27, True),
        fill=(35, 80, 125, 255)
    )

    # Formula
    formula_font = get_font(31, True)

    formula_lines = wrap_text(
        draw,
        safe_text(formula),
        formula_font,
        850
    )

    y = 1480

    for line in formula_lines[:2]:
        draw.text(
            (85, y),
            line,
            font=formula_font,
            fill=(20, 30, 40, 255)
        )
        y += 45

    # Result
    draw.rounded_rectangle(
        (75, 1580, 1005, 1670),
        radius=22,
        fill=(35, 130, 80, 245)
    )

    result_font = get_font(36, True)

    result_text = f"✓ Result: {safe_text(result)}"

    draw.text(
        (105, 1600),
        result_text,
        font=result_font,
        fill=(255, 255, 255, 255)
    )

    img.save(output)


# ============================================================
# CREATE CTA / BRANDING OVERLAY
# ============================================================

def create_cta(logo_path, output):
    img = Image.new(
        "RGBA",
        (WIDTH, HEIGHT),
        (0, 0, 0, 0)
    )

    draw = ImageDraw.Draw(img)

    # CTA card
    draw.rounded_rectangle(
        (45, 95, 1035, 470),
        radius=38,
        fill=(18, 25, 38, 250)
    )

    # Logo
    logo_file = Path(logo_path)

    if logo_file.exists():
        try:
            logo = Image.open(logo_file).convert("RGBA")

            logo.thumbnail((175, 175))

            logo_x = 90
            logo_y = 160

            img.alpha_composite(
                logo,
                (logo_x, logo_y)
            )

        except Exception as exc:
            print(f"WARNING: Could not load logo: {exc}")

    # CTA text
    draw.text(
        (300, 150),
        "FOLLOW FOR MORE",
        font=get_font(37, True),
        fill=(255, 255, 255, 255)
    )

    draw.text(
        (300, 215),
        "Excel & Workplace Tips",
        font=get_font(29, False),
        fill=(215, 225, 235, 255)
    )

    draw.text(
        (300, 275),
        "Learn Verse",
        font=get_font(48, True),
        fill=(255, 255, 255, 255)
    )

    draw.text(
        (300, 345),
        "@LearnVerse9556",
        font=get_font(27, False),
        fill=(170, 195, 220, 255)
    )

    img.save(output)


# ============================================================
# RUN FFMPEG
# ============================================================

def run_ffmpeg(raw_video, voice, output, hook, formula, cta):
    raw_video = Path(raw_video)
    voice = Path(voice)
    output = Path(output)

    print()
    print("========================================")
    print("LEARN VERSE VIDEO RENDER")
    print("========================================")
    print(f"Raw video : {raw_video}")
    print(f"Voice     : {voice}")
    print(f"Output    : {output}")
    print()

    # --------------------------------------------------------
    # Validate inputs
    # --------------------------------------------------------

    if not raw_video.exists():
        raise FileNotFoundError(
            f"Raw Calc recording does not exist: {raw_video}"
        )

    if raw_video.stat().st_size < 10000:
        raise RuntimeError(
            f"Raw Calc recording is too small or empty: "
            f"{raw_video.stat().st_size} bytes"
        )

    if not voice.exists():
        raise FileNotFoundError(
            f"Voice file does not exist: {voice}"
        )

    if voice.stat().st_size < 1000:
        raise RuntimeError(
            f"Voice file is too small: {voice.stat().st_size} bytes"
        )

    for overlay in [hook, formula, cta]:
        if not Path(overlay).exists():
            raise FileNotFoundError(
                f"Overlay file missing: {overlay}"
            )

    # --------------------------------------------------------
    # FFmpeg filter
    # --------------------------------------------------------

    filter_complex = (
        # Real Calc recording
        "[0:v]"
        "scale="
        f"{CALC_W}:{CALC_H}:"
        "force_original_aspect_ratio=decrease,"
        f"pad={CALC_W}:{CALC_H}:(ow-iw)/2:(oh-ih)/2:"
        "color=white,"
        "setsar=1"
        "[calc];"

        # Vertical background
        f"color=c=white:s={WIDTH}x{HEIGHT}:r={FPS}:d={DURATION}"
        "[bg];"

        # Put Calc recording in center
        "[bg][calc]"
        f"overlay={CALC_X}:{CALC_Y}"
        "[base];"

        # Hook
        "[base][1:v]"
        "overlay=0:0:enable='between(t,0,4.5)'"
        "[v1];"

        # Formula
        "[v1][2:v]"
        "overlay=0:0:enable='between(t,3.0,11.0)'"
        "[v2];"

        # CTA
        "[v2][3:v]"
        "overlay=0:0:enable='between(t,10.5,14)'"
        "[vout];"

        # Voice normalization
        "[4:a]"
        "aresample=async=1,"
        "loudnorm=I=-16:TP=-1.5:LRA=11"
        "[aout]"
    )

    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-y",

        # Real Calc recording
        "-stream_loop",
        "-1",
        "-i",
        str(raw_video),

        # Hook
        "-loop",
        "1",
        "-i",
        str(hook),

        # Formula
        "-loop",
        "1",
        "-i",
        str(formula),

        # CTA
        "-loop",
        "1",
        "-i",
        str(cta),

        # Voice
        "-i",
        str(voice),

        "-filter_complex",
        filter_complex,

        "-map",
        "[vout]",

        "-map",
        "[aout]",

        "-t",
        str(DURATION),

        # Video
        "-c:v",
        "libx264",

        "-preset",
        "veryfast",

        "-crf",
        "22",

        "-pix_fmt",
        "yuv420p",

        "-r",
        str(FPS),

        # Audio
        "-c:a",
        "aac",

        "-b:a",
        "128k",

        "-ar",
        "48000",

        # Streaming-friendly MP4
        "-movflags",
        "+faststart",

        str(output)
    ]

    print("Running FFmpeg...")
    print()

    result = subprocess.run(
        cmd,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT
    )

    print(result.stdout)

    if result.returncode != 0:
        raise RuntimeError(
            f"FFmpeg failed with exit code {result.returncode}"
        )

    # --------------------------------------------------------
    # Validate final output
    # --------------------------------------------------------

    if not output.exists():
        raise RuntimeError(
            "FFmpeg finished without creating the output MP4."
        )

    if output.stat().st_size < 50000:
        raise RuntimeError(
            f"Output MP4 is suspiciously small: "
            f"{output.stat().st_size} bytes"
        )

    print()
    print("========================================")
    print("VIDEO CREATED SUCCESSFULLY")
    print("========================================")
    print(f"File: {output}")
    print(f"Size: {output.stat().st_size:,} bytes")
    print()


# ============================================================
# MAIN
# ============================================================

def main():
    if len(sys.argv) != 9:
        print(
            "Usage:\n"
            "python render_video.py "
            "raw.mp4 voice.mp3 output.mp4 "
            "\"problem\" \"title\" \"formula\" \"result\" logo.png"
        )

        raise SystemExit(2)

    raw_video = Path(sys.argv[1])
    voice = Path(sys.argv[2])
    output = Path(sys.argv[3])

    problem = sys.argv[4]
    title = sys.argv[5]
    formula = sys.argv[6]
    result = sys.argv[7]
    logo = Path(sys.argv[8])

    # Temporary overlay directory
    with tempfile.TemporaryDirectory(
        prefix="learnverse_render_"
    ) as temp_dir:

        temp = Path(temp_dir)

        hook = temp / "hook.png"
        formula_overlay = temp / "formula.png"
        cta = temp / "cta.png"

        print("Creating overlay graphics...")

        create_hook(
            problem,
            title,
            hook
        )

        create_formula_overlay(
            formula,
            result,
            formula_overlay
        )

        create_cta(
            logo,
            cta
        )

        print("Overlay graphics created.")

        run_ffmpeg(
            raw_video,
            voice,
            output,
            hook,
            formula_overlay,
            cta
        )


if __name__ == "__main__":
    main()
