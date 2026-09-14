#!/usr/bin/env python3
from __future__ import annotations
import argparse
import json
import os
import subprocess
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl.formula.translate import Translator

from formula_library import FORMULAS, get_daily_formulas

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"

def sh(cmd, env=None):
    cmd = [str(x) for x in cmd]
    print("$", " ".join(cmd))
    result = subprocess.run(cmd, env=env, check=False)
    if result.returncode:
        raise RuntimeError(f"Command failed ({result.returncode}): {' '.join(cmd)}")

def translate_formula(formula):
    if not formula or not formula.startswith("="):
        return formula
    return Translator(formula, origin="A2").translate_formula("A5")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=None)
    parser.add_argument("--count", type=int, default=3)
    args = parser.parse_args()

    run_date = datetime.strptime(args.date, "%Y-%m-%d").date() if args.date else date.today()
    if not 1 <= args.count <= len(FORMULAS):
        raise SystemExit(f"count must be 1-{len(FORMULAS)}")

    display = os.environ.get("DISPLAY")
    if not display:
        raise RuntimeError("DISPLAY is not set. The workflow must run generate_daily.py under xvfb-run.")

    OUT.mkdir(exist_ok=True)
    for p in OUT.iterdir():
        if p.is_file():
            p.unlink()

    selected = get_daily_formulas(run_date, args.count)
    indexes = [FORMULAS.index(item) for item in selected]
    manifest = {"date": run_date.isoformat(), "videos": []}

    for n, (idx, item) in enumerate(zip(indexes, selected), 1):
        stem = f"excel_{run_date.isoformat()}_{n:02d}_{item['id']}"
        xlsx = OUT / f"{stem}.xlsx"
        raw = Path("/tmp") / f"{stem}_calc.mp4"
        voice = OUT / f"{stem}.mp3"
        final = OUT / f"{stem}.mp4"

        for p in (xlsx, raw, voice, final):
            p.unlink(missing_ok=True)

        formula = translate_formula(item["formula"])
        print(f"\n=== VIDEO {n}: {item['title']} ===")
        print(f"Formula: {formula}")

        sh([sys.executable, ROOT / "make_workbook.py", idx, xlsx])
        sh([sys.executable, ROOT / "make_voice.py", item["voice"], voice])

        env = os.environ.copy()
        env.update({
            "DISPLAY": display,
            "EXCEL_FORMULA": formula,
            "FORMULA_COL": str(len(item["headers"]) + 2),
            "DEMO_VALUE": str(item["rows"][0][0]) if item["rows"] else "Learn Verse",
            "RECORD_SECONDS": "15",
            "CAPTURE_WIDTH": "1365",
            "CAPTURE_HEIGHT": "900",
            "CAPTURE_FPS": "30",
        })

        sh([sys.executable, ROOT / "record_calc.py", xlsx, raw], env=env)

        sh([
            sys.executable, ROOT / "render_video.py",
            raw, voice, final,
            item["problem"], item["title"], formula, item["result"],
            ROOT / "assets" / "learnverse_logo.png",
        ])

        manifest["videos"].append({
            "file": final.name,
            "workbook": xlsx.name,
            "formula_id": item["id"],
            "problem": item["problem"],
            "title": item["title"],
            "formula": formula,
            "result": item["result"],
            "voice": item["voice"],
        })

        voice.unlink(missing_ok=True)
        raw.unlink(missing_ok=True)

    (OUT / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(json.dumps(manifest, indent=2))

if __name__ == "__main__":
    main()
