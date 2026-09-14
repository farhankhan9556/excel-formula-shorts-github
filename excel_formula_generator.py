from __future__ import annotations

import argparse
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from formula_library import get_daily_formulas

W, H = 1080, 1920
FPS = 15
DURATION = 12.0

BG = (245, 247, 250)
GRID = (205, 211, 220)
TEXT = (25, 31, 38)
MUTED = (92, 101, 112)
HEADER = (31, 41, 55)
WHITE = (255, 255, 255)
GREEN = (30, 120, 75)
LIGHT_GREEN = (226, 244, 234)
BLUE = (45, 92, 170)
LIGHT_BLUE = (229, 238, 252)
YELLOW = (255, 242, 190)
DARK = (17, 24, 39)

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"


def font(size: int, bold=False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf" if bold
        else "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return ImageFont.truetype(p, size)
    return ImageFont.load_default()


F_TITLE = font(54, True)
F_SUB = font(30, False)
F_SHEET = font(27, False)
F_SHEET_BOLD = font(27, True)
F_FORMULA = font(31, True)
F_SMALL = font(24, False)
F_RESULT = font(52, True)
F_END = font(38, True)


def rounded(draw, xy, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def fit_text(draw, text, fnt, max_width):
    """Reduce font size until text fits."""
    size = fnt.size if hasattr(fnt, "size") else 30
    while size > 18:
        ff = font(size, size >= 40)
        if draw.textbbox((0, 0), text, font=ff)[2] <= max_width:
            return ff
        size -= 2
    return font(18)


def cell_rect(x, y, cw, ch, row, col):
    return (x + col * cw, y + row * ch, x + (col + 1) * cw, y + (row + 1) * ch)


def draw_centered(draw, rect, text, fnt, fill=TEXT):
    x1, y1, x2, y2 = rect
    bb = draw.textbbox((0, 0), text, font=fnt)
    tw, th = bb[2] - bb[0], bb[3] - bb[1]
    draw.text(((x1 + x2 - tw) / 2, (y1 + y2 - th) / 2 - 2), text, font=fnt, fill=fill)


def build_frame(item, t):
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)

    # Header
    d.text((60, 52), "EXCEL QUICK TIP", font=F_SUB, fill=BLUE)
    title = item["title"]
    d.text((60, 94), title, font=F_TITLE, fill=TEXT)

    # Progress bar
    d.rounded_rectangle((60, 175, 1020, 184), radius=5, fill=GRID)
    progress = min(max(t / DURATION, 0), 1)
    d.rounded_rectangle((60, 175, 60 + int(960 * progress), 184), radius=5, fill=BLUE)

    # Excel window
    sx, sy, sw, sh = 45, 230, 990, 930
    rounded(d, (sx, sy, sx + sw, sy + sh), 20, WHITE, GRID, 2)

    # Toolbar
    rounded(d, (sx + 20, sy + 20, sx + sw - 20, sy + 78), 12, (242, 244, 247))
    d.text((sx + 40, sy + 34), "Book1.xlsx", font=F_SMALL, fill=MUTED)
    d.text((sx + 230, sy + 34), "fx", font=F_FORMULA, fill=GREEN)

    # Formula bar
    fb_y = sy + 88
    rounded(d, (sx + 20, fb_y, sx + sw - 20, fb_y + 66), 10, WHITE, GRID, 2)

    formula_t = max(0.0, min(1.0, (t - 5.0) / 2.0))
    shown_formula = item["formula"][: max(0, int(len(item["formula"]) * formula_t))]
    if t >= 5.0:
        d.text((sx + 38, fb_y + 17), shown_formula, font=F_FORMULA, fill=TEXT)

    # Worksheet
    table_x = sx + 30
    table_y = sy + 185
    table_w = sw - 60
    cols = len(item["headers"]) + 1
    rows = len(item["rows"]) + 1
    cw = table_w / cols
    ch = 82

    # Row numbers and cells
    for r in range(rows):
        for c in range(cols):
            rect = cell_rect(table_x, table_y, cw, ch, r, c)
            fill = WHITE
            if c == 0:
                fill = (242, 244, 247)
            if r == 0 and c > 0:
                fill = (232, 239, 247)
            d.rectangle(rect, fill=fill, outline=GRID, width=2)

            if r == 0 and c > 0:
                txt = item["headers"][c - 1]
                fnt = F_SHEET_BOLD
            elif r > 0 and c > 0:
                txt = item["rows"][r - 1][c - 1] if c - 1 < len(item["rows"][r - 1]) else ""
                fnt = F_SHEET
            else:
                txt = str(r) if c == 0 else ""
                fnt = F_SMALL

            if txt:
                maxw = int(cw - 20)
                fnt2 = fit_text(d, str(txt), fnt, maxw)
                draw_centered(d, rect, str(txt), fnt2, TEXT if c > 0 else MUTED)

    # Formula cell coordinates.
    col_letter = item["formula_cell"][0]
    row_num = int(item["formula_cell"][1:])
    col_index = ord(col_letter.upper()) - ord("A") + 1

    # Add columns if formula cell is beyond displayed table.
    if col_index >= cols:
        # Draw a dedicated formula/result panel instead.
        formula_rect = (table_x, table_y + rows * ch + 30, table_x + table_w, table_y + rows * ch + 160)
    else:
        formula_rect = cell_rect(table_x, table_y, cw, ch, row_num, col_index)

    # Typing / selection states
    if t < 3.0:
        status = "Step 1  •  Enter your data"
        status_fill = BLUE
    elif t < 5.0:
        status = "Step 2  •  Select the result cell"
        status_fill = BLUE
    elif t < 8.0:
        status = "Step 3  •  Type the formula"
        status_fill = GREEN
        if col_index < cols:
            d.rectangle(formula_rect, fill=LIGHT_BLUE, outline=BLUE, width=4)
            shown = shown_formula
            draw_centered(d, formula_rect, shown if shown else "|", fit_text(d, shown if shown else "|", F_SHEET, int(cw - 16)), TEXT)
    elif t < 10.2:
        status = "Step 4  •  Press Enter"
        status_fill = GREEN
        if col_index < cols:
            d.rectangle(formula_rect, fill=LIGHT_GREEN, outline=GREEN, width=5)
            draw_centered(d, formula_rect, item["result"], fit_text(d, item["result"], F_SHEET_BOLD, int(cw - 16)), GREEN)
    else:
        status = "Result  ✓"
        status_fill = GREEN
        if col_index < cols:
            d.rectangle(formula_rect, fill=LIGHT_GREEN, outline=GREEN, width=6)
            draw_centered(d, formula_rect, item["result"], fit_text(d, item["result"], F_SHEET_BOLD, int(cw - 16)), GREEN)

    # Status chip
    chip_y = sy + sh - 90
    rounded(d, (sx + 30, chip_y, sx + sw - 30, chip_y + 55), 15, LIGHT_GREEN if status_fill == GREEN else LIGHT_BLUE)
    d.text((sx + 50, chip_y + 13), status, font=F_SUB, fill=status_fill)

    # Formula explanation panel
    py = 1210
    rounded(d, (45, py, 1035, 1470), 22, WHITE, GRID, 2)
    d.text((75, py + 35), "FORMULA", font=F_SMALL, fill=MUTED)

    formula_display = item["formula"]
    ff = fit_text(d, formula_display, F_FORMULA, 900)
    d.text((75, py + 82), formula_display, font=ff, fill=DARK)

    explanation = item["explanation"]
    ef = fit_text(d, explanation, F_SUB, 900)
    d.text((75, py + 145), explanation, font=ef, fill=MUTED)

    # Result badge
    rounded(d, (75, py + 195, 1005, py + 240), 12, LIGHT_GREEN)
    result_text = f"Outcome: {item['result']}"
    d.text((95, py + 204), result_text, font=F_SUB, fill=GREEN)

    # Ending
    if t >= 10.2:
        alpha = min(1.0, (t - 10.2) / 0.7)
        # simple opaque card; timing remains clean and readable
        rounded(d, (95, 1515, 985, 1740), 25, DARK)
        d.text((145, 1560), "SAVE THIS FORMULA", font=F_END, fill=WHITE)
        d.text((145, 1630), "Useful at work • Excel shortcut", font=F_SUB, fill=WHITE)

    # Footer
    d.text((60, 1810), "Excel Formula Shorts", font=F_SMALL, fill=MUTED)
    d.text((W - 350, 1810), "Follow for more tips", font=F_SMALL, fill=MUTED)

    return img


def require_ffmpeg():
    if shutil.which("ffmpeg") is None:
        raise RuntimeError("FFmpeg was not found. Install FFmpeg or use the GitHub Actions workflow.")


def render_video(item, output_path: Path):
    require_ffmpeg()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        "ffmpeg", "-y",
        "-f", "rawvideo",
        "-vcodec", "rawvideo",
        "-pix_fmt", "rgb24",
        "-s", f"{W}x{H}",
        "-r", str(FPS),
        "-i", "-",
        "-an",
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "24",
        "-pix_fmt", "yuv420p",
        "-movflags", "+faststart",
        str(output_path),
    ]

    proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    try:
        total = int(DURATION * FPS)
        for frame_no in range(total):
            t = frame_no / FPS
            frame = build_frame(item, t)
            proc.stdin.write(frame.tobytes())
        proc.stdin.close()
        stderr = proc.stderr.read().decode("utf-8", errors="replace")
        rc = proc.wait()
        if rc != 0:
            raise RuntimeError(f"FFmpeg failed for {output_path.name}:\n{stderr[-4000:]}")
    except Exception:
        try:
            proc.kill()
        except Exception:
            pass
        raise


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD; defaults to today's date")
    parser.add_argument("--count", type=int, default=3)
    args = parser.parse_args()

    if args.date:
        run_date = datetime.strptime(args.date, "%Y-%m-%d").date()
    else:
        run_date = date.today()

    OUT.mkdir(exist_ok=True)
    # Prevent old local files from being mixed into a run.
    for p in OUT.glob("*.mp4"):
        p.unlink()
    for p in OUT.glob("*.json"):
        p.unlink()

    selected = get_daily_formulas(run_date, args.count)
    manifest = {
        "date": run_date.isoformat(),
        "videos": []
    }

    for idx, item in enumerate(selected, start=1):
        safe_id = item["id"]
        filename = f"excel_formula_{run_date.isoformat()}_{idx:02d}_{safe_id}.mp4"
        path = OUT / filename

        print(f"Rendering {idx}/{len(selected)}: {item['title']}")
        render_video(item, path)

        manifest["videos"].append({
            "file": filename,
            "formula_id": item["id"],
            "title": item["title"],
            "formula": item["formula"],
            "result": item["result"],
            "explanation": item["explanation"],
            "duration_seconds": DURATION,
            "fps": FPS,
            "resolution": f"{W}x{H}",
        })

    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps(manifest, indent=2))
    print(f"\nCreated {len(selected)} videos in {OUT}")


if __name__ == "__main__":
    main()
