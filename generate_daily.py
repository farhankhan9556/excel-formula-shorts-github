from __future__ import annotations
import argparse, json, os, shutil, subprocess, sys, time
from datetime import date, datetime
from pathlib import Path

from formula_library import FORMULAS, get_daily_formulas

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"

def sh(cmd, env=None):
    print("$", " ".join(map(str, cmd)))
    subprocess.run([str(x) for x in cmd], check=True, env=env)

def start_xvfb():
    subprocess.run(["pkill", "-f", "Xvfb :99"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.Popen(["Xvfb", ":99", "-screen", "0", "1365x900x24", "-ac"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    time.sleep(2)

def record_calc(xlsx, raw_video, item):
    env = os.environ.copy()
    env["DISPLAY"] = ":99"
    env["EXCEL_FORMULA"] = item["formula"]
    env["FORMULA_COL"] = str(len(item["headers"]) + 2)
    env["RECORD_SECONDS"] = "11"

    # Record X display for 9 seconds in parallel with Calc automation.
    recorder = subprocess.Popen([
        "ffmpeg","-y","-f","x11grab","-video_size","1365x900","-framerate","30",
        "-i",":99.0","-c:v","libx264","-preset","veryfast","-crf","25",
        "-pix_fmt","yuv420p",str(raw_video)
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env)

    time.sleep(1)
    sh([sys.executable, ROOT/"record_calc.py", xlsx, raw_video], env=env)
    try:
        recorder.terminate()
        recorder.wait(timeout=5)
    except Exception:
        recorder.kill()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None)
    ap.add_argument("--count", type=int, default=3)
    args = ap.parse_args()

    run_date = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()
    OUT.mkdir(exist_ok=True)

    for p in OUT.iterdir():
        if p.is_file():
            p.unlink()

    selected = get_daily_formulas(run_date, args.count)
    # Determine indexes in the master library.
    indexes = [FORMULAS.index(item) for item in selected]
    manifest = {"date": run_date.isoformat(), "videos": []}

    start_xvfb()

    for n, (idx, item) in enumerate(zip(indexes, selected), 1):
        stem = f"excel_{run_date.isoformat()}_{n:02d}_{item['id']}"
        xlsx = OUT / f"{stem}.xlsx"
        raw = Path("/tmp") / f"{stem}_calc.mp4"
        voice = OUT / f"{stem}.mp3"
        final = OUT / f"{stem}.mp4"

        print(f"\n=== VIDEO {n}: {item['title']} ===")
        sh([sys.executable, ROOT/"make_workbook.py", idx, xlsx])
        sh([sys.executable, ROOT/"make_voice.py", item["voice"], voice])
        record_calc(xlsx, raw, item)
        sh([
            sys.executable, ROOT/"render_video.py",
            raw, voice, final,
            item["problem"], item["title"], item["formula"], item["result"],
            ROOT/"assets/learnverse_logo.png"
        ])

        # Remove intermediate audio; keep the actual workbook.
        voice.unlink(missing_ok=True)
        raw.unlink(missing_ok=True)

        manifest["videos"].append({
            "file": final.name,
            "workbook": xlsx.name,
            "formula_id": item["id"],
            "problem": item["problem"],
            "title": item["title"],
            "formula": item["formula"],
            "result": item["result"],
            "voice": item["voice"]
        })

    (OUT/"manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps(manifest, indent=2))

if __name__ == "__main__":
    main()
