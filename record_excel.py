from __future__ import annotations

import subprocess
from pathlib import Path

from excel_actions import focus_excel


def record_excel(excel, out, duration=30):
    """Start a robust full-desktop GDI recording and return the Popen process.

    We deliberately do NOT crop to the Excel window here. Window-coordinate cropping
    was the source of the GitHub Actions FFmpeg failure. Excel is maximized, so the
    final caption stage performs the controlled A:G crop instead.
    """
    out = Path(out).resolve()
    focus_excel(excel)

    log_file = out.with_name(out.stem + "_ffmpeg.log")
    log_handle = open(log_file, "w", encoding="utf-8")

    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel", "warning",
        "-y",
        "-f", "gdigrab",
        "-framerate", "30",
        "-draw_mouse", "1",
        "-i", "desktop",
        "-t", str(duration),
        "-an",
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "18",
        "-pix_fmt", "yuv420p",
        str(out),
    ]

    print("Desktop FFmpeg command:")
    print(" ".join(cmd))
    print(f"FFmpeg log: {log_file}")

    try:
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.DEVNULL,
            stdout=log_handle,
            stderr=subprocess.STDOUT,
            creationflags=getattr(
                subprocess,
                "CREATE_NEW_PROCESS_GROUP",
                0,
            ),
        )
    except Exception:
        log_handle.close()
        raise

    process._learnverse_log_handle = log_handle
    return process
