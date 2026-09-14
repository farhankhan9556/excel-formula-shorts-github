#!/usr/bin/env python3
from __future__ import annotations
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

DISPLAY = os.environ.get("DISPLAY", ":99")
WIDTH = int(os.environ.get("CAPTURE_WIDTH", "1365"))
HEIGHT = int(os.environ.get("CAPTURE_HEIGHT", "900"))
FPS = int(os.environ.get("CAPTURE_FPS", "30"))
SECONDS = int(os.environ.get("RECORD_SECONDS", "15"))

def run(cmd, check=False, capture=False):
    return subprocess.run(
        cmd,
        check=check,
        stdout=subprocess.PIPE if capture else subprocess.DEVNULL,
        stderr=subprocess.STDOUT if capture else subprocess.DEVNULL,
        text=True,
    )

def output(cmd):
    r = run(cmd, capture=True)
    return r.returncode, r.stdout or ""

def wait_for_calc(timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        rc, text = output(["xdotool", "search", "--onlyvisible", "--class", "libreoffice"])
        if rc == 0 and text.strip():
            return text.strip().splitlines()[0]

        rc, text = output(["xdotool", "search", "--onlyvisible", "--name", ".*"])
        if rc == 0:
            for wid in text.strip().splitlines():
                if not wid.strip():
                    continue
                rc2, name = output(["xdotool", "getwindowname", wid])
                low = name.lower()
                if rc2 == 0 and ("calc" in low or ".xlsx" in low or "excel tip" in low):
                    return wid
        time.sleep(0.5)
    raise RuntimeError(f"LibreOffice Calc window was not detected on {DISPLAY}")

def key(name):
    rc, text = output(["xdotool", "key", "--clearmodifiers", name])
    if rc:
        raise RuntimeError(f"xdotool key failed: {name}\n{text}")

def type_text(text, delay=25):
    rc, out = output(["xdotool", "type", "--clearmodifiers", "--delay", str(delay), text])
    if rc:
        raise RuntimeError(f"xdotool type failed\n{out}")

def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: python record_calc.py workbook.xlsx output.mp4")

    xlsx = Path(sys.argv[1]).resolve()
    video = Path(sys.argv[2]).resolve()
    if not xlsx.exists():
        raise FileNotFoundError(xlsx)

    video.parent.mkdir(parents=True, exist_ok=True)
    video.unlink(missing_ok=True)

    formula = os.environ.get("EXCEL_FORMULA", "").strip()
    formula_col = int(os.environ.get("FORMULA_COL", "5"))
    demo_value = os.environ.get("DEMO_VALUE", "Learn Verse")

    # Confirm the X display is usable before starting LibreOffice.
    rc, out = output(["xdpyinfo", "-display", DISPLAY])
    if rc:
        raise RuntimeError(f"X display {DISPLAY} is not available.\n{out}")

    # Remove only exact stale LibreOffice processes.
    run(["pkill", "-x", "soffice.bin"])
    run(["pkill", "-x", "soffice"])
    time.sleep(1)

    ffmpeg_cmd = [
        "ffmpeg", "-hide_banner", "-loglevel", "warning", "-y",
        "-f", "x11grab",
        "-draw_mouse", "1",
        "-video_size", f"{WIDTH}x{HEIGHT}",
        "-framerate", str(FPS),
        "-i", f"{DISPLAY}.0",
        "-t", str(SECONDS),
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "23",
        "-pix_fmt", "yuv420p",
        str(video),
    ]
    recorder = subprocess.Popen(
        ffmpeg_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )

    calc = None
    profile = Path(tempfile.mkdtemp(prefix="learnverse-lo-"))

    try:
        time.sleep(1)

        env = os.environ.copy()
        env["DISPLAY"] = DISPLAY
        env["SAL_USE_VCLPLUGIN"] = "gen"
        env["SAL_DISABLE_OPENCL"] = "1"

        calc = subprocess.Popen(
            [
                "libreoffice",
                "--nologo",
                "--nodefault",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation=file://{profile}",
                str(xlsx),
            ],
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )

        wid = wait_for_calc()
        run(["xdotool", "windowactivate", "--sync", wid])
        run(["xdotool", "windowraise", wid])
        run(["xdotool", "windowsize", wid, str(WIDTH), str(HEIGHT)])
        run(["xdotool", "windowmove", wid, "0", "0"])
        time.sleep(1)

        # Go to the Live Entry cell and type a harmless demo value.
        key("ctrl+home")
        for _ in range(4):
            key("down")
        for _ in range(formula_col + 2 - 1):
            key("right")
        key("f2")
        key("ctrl+a")
        type_text(demo_value, 35)
        key("Return")
        time.sleep(0.8)

        # Go to the genuine formula cell and re-enter the exact translated formula.
        key("ctrl+home")
        for _ in range(4):
            key("down")
        for _ in range(formula_col - 1):
            key("right")
        key("f2")
        key("ctrl+a")
        if formula:
            type_text(formula, 20)
        key("Return")
        time.sleep(2)

        # Keep the genuine Calc result visible.
        remaining = max(1, SECONDS - 6)
        time.sleep(remaining)

    finally:
        run(["xdotool", "key", "alt+F4"])
        time.sleep(1)
        run(["pkill", "-x", "soffice.bin"])
        run(["pkill", "-x", "soffice"])

        try:
            recorder.wait(timeout=15)
        except subprocess.TimeoutExpired:
            recorder.terminate()
            try:
                recorder.wait(timeout=5)
            except subprocess.TimeoutExpired:
                recorder.kill()
                recorder.wait()

    if not video.exists() or video.stat().st_size < 10000:
        ffout = ""
        if recorder.stdout:
            try:
                ffout = recorder.stdout.read()
            except Exception:
                pass
        calc_err = ""
        if calc is not None and calc.stderr:
            try:
                calc_err = calc.stderr.read()
            except Exception:
                pass
        raise RuntimeError(
            f"Calc recording was not created correctly: {video}\n"
            f"FFmpeg output:\n{ffout}\n"
            f"LibreOffice output:\n{calc_err}"
        )

    probe = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-show_entries", "format=duration,size",
            "-show_entries", "stream=codec_name,width,height",
            "-of", "default=noprint_wrappers=1",
            str(video),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    if probe.returncode:
        raise RuntimeError(f"Recorded MP4 failed ffprobe:\n{probe.stdout}")

    print("Calc recording validated:")
    print(probe.stdout)

if __name__ == "__main__":
    main()
