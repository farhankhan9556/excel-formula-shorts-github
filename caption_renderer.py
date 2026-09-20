from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

W, H = 1080, 1920
FONT = Path(r"C:\Windows\Fonts\arialbd.ttf")
if not FONT.exists():
    FONT = Path(r"C:\Windows\Fonts\segoeuib.ttf")


def F(size):
    return ImageFont.truetype(str(FONT), size)


def wrap(draw, text, font, max_width):
    words = str(text).split()
    lines = []
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines or [""]


def rounded_card(draw, box, radius=28, fill=(18, 25, 35, 230), outline=(255, 255, 255, 70)):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2)


def make_infographic(out, formula_name, formula_text, steps, explanation):
    """Professional transparent lower-third with three stacked cards."""
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)

    # A very light transparent veil only behind the cards.
    d.rounded_rectangle(
        (36, 1060, 1044, 1888),
        radius=36,
        fill=(8, 13, 20, 90),
        outline=(255, 255, 255, 45),
        width=2,
    )

    title_font = F(34)
    body_font = F(27)
    small_font = F(24)
    formula_font = F(25)

    # Card 1: Today’s Steps
    y = 1085
    rounded_card(d, (58, y, 1022, y + 225))
    d.text((88, y + 24), "TODAY'S STEPS", font=title_font, fill=(255, 255, 255, 255))
    step_text = []
    for i, s in enumerate(steps[:3], 1):
        step_text.append(f"{i}. {s}")
    yy = y + 78
    for line in step_text:
        lines = wrap(d, line, body_font, 890)
        for sub in lines[:2]:
            d.text((95, yy), sub, font=body_font, fill=(245, 248, 252, 255))
            yy += 34
        yy += 4

    # Card 2: Formula Used
    y = 1335
    rounded_card(d, (58, y, 1022, y + 305))
    d.text((88, y + 22), "FORMULA USED", font=title_font, fill=(255, 255, 255, 255))
    d.text((88, y + 72), str(formula_name), font=title_font, fill=(255, 242, 160, 255))
    formula_lines = wrap(d, str(formula_text), formula_font, 875)
    yy = y + 119
    for line in formula_lines[:2]:
        d.text((88, yy), line, font=formula_font, fill=(255, 255, 255, 255))
        yy += 34
    exp_lines = wrap(d, str(explanation), small_font, 875)
    yy += 8
    for line in exp_lines[:3]:
        d.text((88, yy), line, font=small_font, fill=(220, 228, 238, 255))
        yy += 30

    # Card 3: YouTube branding
    y = 1660
    rounded_card(d, (58, y, 1022, y + 200))
    d.text((88, y + 24), "WATCH ON YOUTUBE", font=title_font, fill=(255, 255, 255, 255))
    d.text((88, y + 82), "Learn Verse", font=title_font, fill=(255, 255, 255, 255))
    d.text((88, y + 132), "@LearnVerse9556  •  Excel tips every day", font=small_font, fill=(225, 232, 240, 255))

    im.save(out, "PNG")


def make_card(text, out):
    # Backward compatibility for older callers.
    im = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    font = F(56)
    y = 1450
    for line in str(text).split("\n"):
        bb = d.textbbox((0, 0), line, font=font, stroke_width=3)
        x = (W - (bb[2] - bb[0])) // 2
        d.text((x, y), line, font=font, fill=(255, 255, 255, 255), stroke_width=4, stroke_fill=(0, 0, 0, 255))
        y += 70
    im.save(out, "PNG")
