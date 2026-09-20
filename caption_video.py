from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from caption_renderer import make_infographic


FORMULAS = {
    "if": ("IF", '=IF(C5="Done","YES","NO")', "IF checks a condition and returns one value when true and another when false."),
    "sum": ("SUM", "=SUM(D5:D10)", "SUM adds all numbers in the selected range and returns the total."),
    "product": ("PRODUCT", "=PRODUCT(B5:B6)", "PRODUCT multiplies the selected values together and returns the result."),
    "countif": ("COUNTIF", '=COUNTIF(C5:C10,"Done")', "COUNTIF counts cells that match the condition you specify."),
    "sumif": ("SUMIF", '=SUMIF(C5:C10,"Done",D5:D10)', "SUMIF adds values only when they meet the selected condition."),
    "average": ("AVERAGE", "=AVERAGE(D5:D10)", "AVERAGE calculates the mean of the selected numbers."),
    "max": ("MAX", "=MAX(D5:D10)", "MAX returns the largest number in the selected range."),
    "min": ("MIN", "=MIN(D5:D10)", "MIN returns the smallest number in the selected range."),
    "round": ("ROUND", "=ROUND(D5,2)", "ROUND rounds a number to the number of digits you choose."),
    "left": ("LEFT", "=LEFT(A5,3)", "LEFT returns a chosen number of characters from the beginning of text."),
    "textjoin": ("TEXTJOIN", '=TEXTJOIN(" ",TRUE,A5:B5)', "TEXTJOIN combines text values using the separator you choose."),
    "concatenate": ("TEXTJOIN", '=TEXTJOIN(" - ",TRUE,E10,"LearnVerse")', "TEXTJOIN combines text values using the separator you choose."),
    "xlookup": ("XLOOKUP", '=XLOOKUP(A5,A8:A13,D8:D13,"Not Found")', "XLOOKUP searches for a value and returns the matching result from another range."),
    "checkbox": ("EXCEL CHECKBOXES", "TRUE / FALSE", "Real Excel Form Controls can be linked to cells and used to update task status automatically."),
}


def topic_from(final: Path) -> str:
    m = re.search(r"_([a-z0-9]+)$", final.stem, re.I)
    key = m.group(1).lower() if m else "if"
    return key if key in FORMULAS else "if"


def media_duration(path: Path) -> float:
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(path),
    ]
    p = subprocess.run(cmd, capture_output=True, text=True)
    try:
        value = float(p.stdout.strip())
        if value > 0:
            return value
    except Exception:
        pass
    return 30.0


def main():
    if len(sys.argv) != 5:
        raise SystemExit("Usage: caption_video.py raw.mp4 voice.mp3 final.mp4 steps_json")

    raw = Path(sys.argv[1]).resolve()
    voice = Path(sys.argv[2]).resolve()
    final = Path(sys.argv[3]).resolve()
    steps = json.loads(sys.argv[4])

    topic = topic_from(final)
    formula_name, formula_text, explanation = FORMULAS[topic]

    clean_steps = []
    for item in steps:
        if isinstance(item, list) and len(item) >= 3:
            text = re.sub(r"^\s*\d+\s+", "", str(item[2]).strip())
            if text:
                clean_steps.append(text)

    if len(clean_steps) < 3:
        clean_steps = [
            "Enter the required value.",
            f"Use the {formula_name} formula.",
            "Press Enter and check the result.",
        ]

    voice_duration = media_duration(voice)
    final_duration = min(30.0, max(1.0, voice_duration + 0.15))

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        info = td / "learnverse_cards.png"
        make_infographic(info, formula_name, formula_text, clean_steps, explanation)

        # Raw recording is the full desktop. Crop from the left so the final Short
        # contains the Excel A:G area after restore_full_excel_view() selects A1.
        filter_complex = (
            "[0:v]"
            "crop=900:ih:0:0,"
            "scale=1080:1092:flags=lanczos,"
            "setsar=1,"
            "pad=1080:1920:0:0:color=white"
            "[excel];"
            "[2:v]format=rgba[cards];"
            "[excel][cards]overlay=0:0:format=auto[video];"
            "[1:a]"
            "highpass=f=80,"
            "lowpass=f=15000,"
            "equalizer=f=1800:t=q:w=1:g=1.5,"
            "equalizer=f=3000:t=q:w=1:g=2,"
            "acompressor=threshold=-19dB:ratio=2.5:attack=10:release=140:makeup=2,"
            "loudnorm=I=-15:TP=-1.5:LRA=6,"
            "aresample=48000"
            "[audio]"
        )

        cmd = [
            "ffmpeg", "-hide_banner", "-loglevel", "warning", "-y",
            "-i", str(raw),
            "-i", str(voice),
            "-loop", "1", "-i", str(info),
            "-filter_complex", filter_complex,
            "-map", "[video]",
            "-map", "[audio]",
            "-t", f"{final_duration:.3f}",
            "-c:v", "libx264",
            "-preset", "veryfast",
            "-crf", "18",
            "-pix_fmt", "yuv420p",
            "-r", "30",
            "-c:a", "aac",
            "-b:a", "160k",
            "-ar", "48000",
            "-ac", "2",
            "-movflags", "+faststart",
            str(final),
        ]

        p = subprocess.run(cmd, capture_output=True, text=True)
        if p.returncode != 0:
            raise RuntimeError(p.stderr[-6000:])

    if not final.exists() or final.stat().st_size < 50000:
        raise RuntimeError("Final video was not created correctly.")

    print(f"FINAL VIDEO CREATED: {final}")
    print(f"VOICE DURATION: {voice_duration:.2f}s")
    print(f"FINAL VIDEO DURATION: {final_duration:.2f}s")
    print("EXCEL VIEW: A:G")


if __name__ == "__main__":
    main()
