from __future__ import annotations

import argparse
import asyncio
import datetime
import json
import subprocess
import time
from pathlib import Path

from make_excel_demo import build_workbook
from excel_actions import type_text, type_formula, restore_full_excel_view
from record_excel import record_excel
from make_voice import make as make_voice


ROOT = Path(__file__).resolve().parent
OUT = ROOT / "output"


def clean_output():
    OUT.mkdir(parents=True, exist_ok=True)

    for p in OUT.iterdir():
        if p.is_file():
            try:
                p.unlink()
            except Exception as exc:
                print(f"WARNING: Could not delete {p}: {exc}")


def find_cell(sheet, value):
    used = sheet.UsedRange
    target = str(value).strip().upper()

    for r in range(
        int(used.Row),
        int(used.Row + used.Rows.Count)
    ):
        for c in range(
            int(used.Column),
            int(used.Column + used.Columns.Count)
        ):
            try:
                v = sheet.Cells(r, c).Value
                if v is not None:
                    if str(v).strip().upper() == target:
                        return r, c
            except Exception:
                pass

    return None


def perform_formula_demo(excel, book, video):
    sheet = book.Worksheets("Learn Verse")
    sheet.Activate()

    input_row = 16
    input_col = 5

    result_row = 16
    result_col = 7

    type_text(
        excel,
        sheet,
        input_row,
        input_col,
        video.get("input", ""),
        zoom=140,
    )

    time.sleep(1.0)

    formula = str(video.get("formula", ""))

    if video.get("id") == "countif":
        formula = "=COUNTIF(C5:C10,E16)"

    type_formula(
        excel,
        sheet,
        result_row,
        result_col,
        formula,
        zoom=140,
    )

    time.sleep(2.0)

    restore_full_excel_view(
        excel,
        sheet,
        zoom=80,
    )

    time.sleep(5.0)


def perform_actions(excel, book, video):
    if video.get("type") == "formula":
        perform_formula_demo(
            excel,
            book,
            video,
        )
    else:
        sheet = book.Worksheets("Learn Verse")

        restore_full_excel_view(
            excel,
            sheet,
            zoom=80,
        )

        from excel_actions import click_cell

        for row in video.get(
            "checkbox_rows",
            [4, 5, 7],
        ):
            click_cell(
                excel,
                sheet,
                row,
                3,
            )
            time.sleep(0.9)


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--date",
        default="",
    )

    parser.add_argument(
        "--count",
        type=int,
        default=3,
    )

    args = parser.parse_args()

    run_date = (
        datetime.datetime.strptime(
            args.date,
            "%Y-%m-%d",
        ).date()
        if args.date
        else datetime.date.today()
    )

    count = max(
        1,
        min(
            3,
            args.count,
        ),
    )

    clean_output()

    from formula_library import get_daily_videos

    videos = get_daily_videos(
        run_date,
        count,
    )

    manifest = []

    for index, video in enumerate(
        videos,
        1,
    ):

        stem = (
            f"learnverse_"
            f"{run_date}_"
            f"{index:02d}_"
            f"{video['id']}"
        )

        xlsx = OUT / f"{stem}.xlsx"
        raw = OUT / f"{stem}_raw.mp4"
        voice = OUT / f"{stem}.mp3"
        final = OUT / f"{stem}.mp4"
        metadata = OUT / f"{stem}.json"

        excel = None
        book = None
        recorder = None

        try:

            print("=" * 60)
            print(
                f"GENERATING VIDEO "
                f"{index}/{len(videos)}"
            )
            print(
                f"TOPIC: "
                f"{video.get('title', video.get('id'))}"
            )
            print("=" * 60)

            excel, book = build_workbook(
                video,
                xlsx,
            )

            excel.Visible = True
            excel.DisplayAlerts = False
            excel.ScreenUpdating = True
            excel.WindowState = -4137

            sheet = book.Worksheets(
                "Learn Verse"
            )

            sheet.Activate()

            restore_full_excel_view(
                excel,
                sheet,
                zoom=80,
            )

            time.sleep(2)

            print(
                "Generating clear tutorial voice..."
            )

            asyncio.run(
                make_voice(
                    video["voice"],
                    voice,
                )
            )

            if not voice.exists():
                raise RuntimeError(
                    "Voice file was not created."
                )

            print(
                "Starting Excel recording..."
            )

            recorder = record_excel(
                excel,
                raw,
                30,
            )

            if recorder is None:
                raise RuntimeError(
                    "FFmpeg recorder did not start."
                )

            print(
                f"FFmpeg process started. "
                f"PID: {recorder.pid}"
            )

            time.sleep(2)

            print(
                "Entering data with "
                "close-up cell zoom..."
            )

            perform_actions(
                excel,
                book,
                video,
            )

            print(
                "Showing final full Excel page "
                "at 80%..."
            )

            time.sleep(4)

            book.Save()

            print(
                "Waiting for FFmpeg recording "
                "to finish..."
            )

            try:
                return_code = recorder.wait(
                    timeout=55
                )

            except subprocess.TimeoutExpired:

                print(
                    "ERROR: FFmpeg recording "
                    "timed out."
                )

                try:
                    recorder.terminate()
                except Exception:
                    pass

                raise RuntimeError(
                    "FFmpeg recording timed out."
                )

            print(
                f"FFmpeg return code: "
                f"{return_code}"
            )

            if return_code != 0:

                log_file = raw.with_name(
                    raw.stem +
                    "_ffmpeg.log"
                )

                if log_file.exists():

                    print(
                        "FFmpeg log:"
                    )

                    try:
                        print(
                            log_file.read_text(
                                encoding="utf-8",
                                errors="replace",
                            )
                        )
                    except Exception:
                        pass

                raise RuntimeError(
                    "FFmpeg Excel recording failed."
                )

            if not raw.exists():

                raise RuntimeError(
                    "FFmpeg finished successfully "
                    "but raw recording file is missing."
                )

            raw_size = raw.stat().st_size

            print(
                f"Raw recording size: "
                f"{raw_size:,} bytes"
            )

            if raw_size < 50000:

                raise RuntimeError(
                    "Raw Excel recording is "
                    "missing or too small."
                )

            print(
                "Creating final 1080x1920 Short..."
            )

            subprocess.run(
                [
                    "python",
                    str(
                        ROOT /
                        "caption_video.py"
                    ),
                    str(raw),
                    str(voice),
                    str(final),
                    json.dumps(
                        video["captions"],
                        ensure_ascii=False,
                    ),
                ],
                check=True,
            )

            if not final.exists():

                raise RuntimeError(
                    "Final MP4 was not created."
                )

            final_size = final.stat().st_size

            if final_size < 50000:

                raise RuntimeError(
                    "Final MP4 is too small."
                )

            payload = {
                "date": str(run_date),
                "video_number": index,
                "id": video["id"],
                "title": video.get(
                    "youtube_title",
                    video.get(
                        "title",
                        "",
                    ),
                ),
                "description": video.get(
                    "youtube_description",
                    "",
                ),
                "hashtags": video.get(
                    "hashtags",
                    [],
                ),
                "upload_slot": video.get(
                    "upload_slot",
                    index,
                ),
                "video_file": final.name,
                "workbook_file": xlsx.name,
            }

            metadata.write_text(
                json.dumps(
                    payload,
                    indent=2,
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            manifest.append(payload)

            print(
                f"FINAL: {final}"
            )

            print(
                f"SIZE: "
                f"{final_size:,} bytes"
            )

            print(
                f"VIDEO {index} COMPLETED"
            )

        finally:

            if recorder is not None:

                try:

                    if recorder.poll() is None:

                        recorder.terminate()

                        try:
                            recorder.wait(
                                timeout=10
                            )
                        except Exception:
                            try:
                                recorder.kill()
                            except Exception:
                                pass

                except Exception:
                    pass

            if recorder is not None:

                try:

                    log_handle = getattr(
                        recorder,
                        "_learnverse_log_handle",
                        None,
                    )

                    if log_handle is not None:
                        log_handle.flush()
                        log_handle.close()

                except Exception:
                    pass

            try:
                raw.unlink(
                    missing_ok=True
                )
            except Exception:
                pass

            try:
                voice.unlink(
                    missing_ok=True
                )
            except Exception:
                pass

            try:

                if book is not None:
                    book.Close(
                        SaveChanges=False
                    )

            except Exception:
                pass

            try:

                if excel is not None:
                    excel.Quit()

            except Exception:
                pass

            time.sleep(1)

    manifest_file = (
        OUT /
        "manifest.json"
    )

    manifest_file.write_text(
        json.dumps(
            {
                "date": str(run_date),
                "videos": manifest,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print("=" * 60)
    print("ALL VIDEOS COMPLETED")
    print("=" * 60)


if __name__ == "__main__":
    main()
