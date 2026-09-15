from __future__ import annotations

import argparse
import asyncio
import datetime as dt
import json
import math
import os
import random
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pyautogui
import win32com.client as win32
import edge_tts


# ============================================================
# LEARN VERSE - EXCEL SHORTS GENERATOR
# Self-contained Windows + GitHub Actions version
# ============================================================

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
OUTPUT.mkdir(parents=True, exist_ok=True)

pyautogui.PAUSE = 0.15
pyautogui.FAILSAFE = False


# ------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------

VOICE_LIST = [
    "en-US-JennyNeural",
    "en-US-AriaNeural",
    "en-US-SaraNeural",
]

VOICE_RATE = "+5%"
VOICE_VOLUME = "+0%"
VOICE_PITCH = "+1Hz"

FPS = 30
MAX_VIDEO_SECONDS = 30

# The Excel window is recorded from the desktop.
# The final video crops the left side and keeps the Excel sheet
# visible in the vertical video.
CAPTURE_WIDTH = 1366
CAPTURE_HEIGHT = 768

# Excel zoom.
EXCEL_ZOOM = 80


# ------------------------------------------------------------
# FORMULA LIBRARY
# ------------------------------------------------------------

FORMULAS = [
    {
        "id": "if",
        "name": "IF",
        "title": "IF Formula",
        "problem": "Automatically show a result based on a condition.",
        "formula": '=IF(C2="Done","Completed","Pending")',
        "explanation": "IF checks a condition and returns one result when it is true and another when it is false.",
        "steps": [
            "Create a Status column.",
            "Select the result cell.",
            'Enter the IF formula.',
            "Press Enter.",
            "The result updates automatically.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Target", "Final"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 100, ""],
            ["Sara", "HR", "Pending", "Medium", "", 100, ""],
            ["Ali", "Finance", "Done", "High", "", 100, ""],
            ["Omar", "IT", "Done", "Low", "", 100, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 100, ""],
            ["Zaid", "Operations", "Done", "High", "", 100, ""],
        ],
        "demo_cell": "E2",
        "formula_cell": "E2",
    },
    {
        "id": "sum",
        "name": "SUM",
        "title": "SUM Formula",
        "problem": "Add a complete range of numbers instantly.",
        "formula": "=SUM(F2:F7)",
        "explanation": "SUM adds all numeric values in the selected range.",
        "steps": [
            "Enter your numbers in a column.",
            "Select the total cell.",
            "Type the SUM formula.",
            "Press Enter.",
            "Excel calculates the total instantly.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Sales", "Total"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 1200, ""],
            ["Sara", "HR", "Pending", "Medium", "", 850, ""],
            ["Ali", "Finance", "Done", "High", "", 1500, ""],
            ["Omar", "IT", "Done", "Low", "", 920, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 760, ""],
            ["Zaid", "Operations", "Done", "High", "", 1350, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "average",
        "name": "AVERAGE",
        "title": "AVERAGE Formula",
        "problem": "Find the average of multiple values.",
        "formula": "=AVERAGE(F2:F7)",
        "explanation": "AVERAGE calculates the arithmetic mean of the selected numbers.",
        "steps": [
            "Enter your values.",
            "Select the result cell.",
            "Type the AVERAGE formula.",
            "Press Enter.",
            "Excel returns the average.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Score", "Average"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 88, ""],
            ["Sara", "HR", "Pending", "Medium", "", 76, ""],
            ["Ali", "Finance", "Done", "High", "", 92, ""],
            ["Omar", "IT", "Done", "Low", "", 81, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 73, ""],
            ["Zaid", "Operations", "Done", "High", "", 95, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "max",
        "name": "MAX",
        "title": "MAX Formula",
        "problem": "Find the highest value in a list.",
        "formula": "=MAX(F2:F7)",
        "explanation": "MAX returns the largest number from the selected range.",
        "steps": [
            "Prepare your numbers.",
            "Select the result cell.",
            "Enter the MAX formula.",
            "Press Enter.",
            "Excel returns the highest value.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Score", "Highest"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 88, ""],
            ["Sara", "HR", "Pending", "Medium", "", 76, ""],
            ["Ali", "Finance", "Done", "High", "", 92, ""],
            ["Omar", "IT", "Done", "Low", "", 81, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 73, ""],
            ["Zaid", "Operations", "Done", "High", "", 95, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "min",
        "name": "MIN",
        "title": "MIN Formula",
        "problem": "Find the lowest value in a list.",
        "formula": "=MIN(F2:F7)",
        "explanation": "MIN returns the smallest number from the selected range.",
        "steps": [
            "Prepare your numbers.",
            "Select the result cell.",
            "Enter the MIN formula.",
            "Press Enter.",
            "Excel returns the lowest value.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Score", "Lowest"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 88, ""],
            ["Sara", "HR", "Pending", "Medium", "", 76, ""],
            ["Ali", "Finance", "Done", "High", "", 92, ""],
            ["Omar", "IT", "Done", "Low", "", 81, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 73, ""],
            ["Zaid", "Operations", "Done", "High", "", 95, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "countif",
        "name": "COUNTIF",
        "title": "COUNTIF Formula",
        "problem": "Count cells that match a condition.",
        "formula": '=COUNTIF(C2:C7,"Done")',
        "explanation": "COUNTIF counts cells that meet the condition you specify.",
        "steps": [
            "Create your status list.",
            "Select the result cell.",
            "Enter the COUNTIF formula.",
            "Press Enter.",
            "Excel counts the matching rows.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Target", "Done Count"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 100, ""],
            ["Sara", "HR", "Pending", "Medium", "", 100, ""],
            ["Ali", "Finance", "Done", "High", "", 100, ""],
            ["Omar", "IT", "Done", "Low", "", 100, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 100, ""],
            ["Zaid", "Operations", "Done", "High", "", 100, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "sumif",
        "name": "SUMIF",
        "title": "SUMIF Formula",
        "problem": "Add values only when a condition is met.",
        "formula": '=SUMIF(C2:C7,"Done",F2:F7)',
        "explanation": "SUMIF adds values from one range when another range matches your condition.",
        "steps": [
            "Create a condition column.",
            "Enter the values to total.",
            "Select the result cell.",
            "Enter the SUMIF formula.",
            "Press Enter.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Sales", "Done Sales"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 1200, ""],
            ["Sara", "HR", "Pending", "Medium", "", 850, ""],
            ["Ali", "Finance", "Done", "High", "", 1500, ""],
            ["Omar", "IT", "Done", "Low", "", 920, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 760, ""],
            ["Zaid", "Operations", "Done", "High", "", 1350, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "round",
        "name": "ROUND",
        "title": "ROUND Formula",
        "problem": "Round numbers to a specific number of decimals.",
        "formula": "=ROUND(F2,2)",
        "explanation": "ROUND reduces a number to the number of decimal places you specify.",
        "steps": [
            "Enter a decimal value.",
            "Select the result cell.",
            "Enter the ROUND formula.",
            "Press Enter.",
            "Excel returns the rounded value.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Amount", "Rounded"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 125.6789, ""],
            ["Sara", "HR", "Pending", "Medium", "", 98.4567, ""],
            ["Ali", "Finance", "Done", "High", "", 245.1299, ""],
            ["Omar", "IT", "Done", "Low", "", 74.9921, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 51.2234, ""],
            ["Zaid", "Operations", "Done", "High", "", 312.9876, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "left",
        "name": "LEFT",
        "title": "LEFT Formula",
        "problem": "Extract characters from the beginning of text.",
        "formula": "=LEFT(A2,3)",
        "explanation": "LEFT returns a specified number of characters from the beginning of text.",
        "steps": [
            "Enter your text.",
            "Select the result cell.",
            "Enter the LEFT formula.",
            "Press Enter.",
            "Excel extracts the characters.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Code", "Extracted"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", "HSE-001", ""],
            ["Sara", "HR", "Pending", "Medium", "", "HR-002", ""],
            ["Ali", "Finance", "Done", "High", "", "FIN-003", ""],
            ["Omar", "IT", "Done", "Low", "", "IT-004", ""],
            ["Hina", "Admin", "Pending", "Medium", "", "ADM-005", ""],
            ["Zaid", "Operations", "Done", "High", "", "OPS-006", ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "product",
        "name": "PRODUCT",
        "title": "PRODUCT Formula",
        "problem": "Multiply numbers together.",
        "formula": "=PRODUCT(F2,F3)",
        "explanation": "PRODUCT multiplies the numbers or cell references you provide.",
        "steps": [
            "Enter your numbers.",
            "Select the result cell.",
            "Enter the PRODUCT formula.",
            "Press Enter.",
            "Excel calculates the multiplication.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Qty", "Price"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 5, 20],
            ["Sara", "HR", "Pending", "Medium", "", 3, 30],
            ["Ali", "Finance", "Done", "High", "", 8, 15],
            ["Omar", "IT", "Done", "Low", "", 4, 25],
            ["Hina", "Admin", "Pending", "Medium", "", 6, 12],
            ["Zaid", "Operations", "Done", "High", "", 10, 18],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "textjoin",
        "name": "TEXTJOIN",
        "title": "TEXTJOIN Formula",
        "problem": "Combine text from multiple cells.",
        "formula": '=TEXTJOIN(" - ",TRUE,A2:C2)',
        "explanation": "TEXTJOIN combines text using a delimiter and can ignore empty cells.",
        "steps": [
            "Enter text in separate cells.",
            "Select the result cell.",
            "Enter the TEXTJOIN formula.",
            "Press Enter.",
            "Excel combines the text.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Code", "Output"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", "001", ""],
            ["Sara", "HR", "Pending", "Medium", "", "002", ""],
            ["Ali", "Finance", "Done", "High", "", "003", ""],
            ["Omar", "IT", "Done", "Low", "", "004", ""],
            ["Hina", "Admin", "Pending", "Medium", "", "005", ""],
            ["Zaid", "Operations", "Done", "High", "", "006", ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
    {
        "id": "xlookup",
        "name": "XLOOKUP",
        "title": "XLOOKUP Formula",
        "problem": "Find related information from another column.",
        "formula": '=XLOOKUP(A2,A2:A7,F2:F7,"Not Found")',
        "explanation": "XLOOKUP searches for a value and returns the matching result.",
        "steps": [
            "Create your lookup table.",
            "Select the result cell.",
            "Enter the XLOOKUP formula.",
            "Press Enter.",
            "Excel returns the matching value.",
        ],
        "headers": ["Employee", "Department", "Status", "Priority", "Result", "Score", "Lookup Result"],
        "rows": [
            ["Ahmed", "HSE", "Done", "High", "", 88, ""],
            ["Sara", "HR", "Pending", "Medium", "", 76, ""],
            ["Ali", "Finance", "Done", "High", "", 92, ""],
            ["Omar", "IT", "Done", "Low", "", 81, ""],
            ["Hina", "Admin", "Pending", "Medium", "", 73, ""],
            ["Zaid", "Operations", "Done", "High", "", 95, ""],
        ],
        "demo_cell": "G2",
        "formula_cell": "G2",
    },
]


# ------------------------------------------------------------
# GENERAL HELPERS
# ------------------------------------------------------------

def log(message: str):
    print(f"[LearnVerse] {message}", flush=True)


def run_command(cmd, timeout=None):
    log("Running: " + " ".join(map(str, cmd)))

    result = subprocess.run(
        [str(x) for x in cmd],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )

    if result.stdout:
        print(result.stdout, flush=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"Command failed with exit code {result.returncode}: "
            + " ".join(map(str, cmd))
        )

    return result


def require_program(name: str):
    if shutil.which(name) is None:
        raise RuntimeError(
            f"{name} was not found in PATH. "
            f"Please install {name} and restart the terminal/runner."
        )


def ffprobe_duration(path: Path) -> float:
    result = subprocess.run(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "default=noprint_wrappers=1:nokey=1",
            str(path),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )

    try:
        return float(result.stdout.strip())
    except Exception:
        return 0.0


def safe_delete(path: Path):
    try:
        if path.exists():
            path.unlink()
    except Exception as exc:
        log(f"Could not delete {path.name}: {exc}")


# ------------------------------------------------------------
# DAILY ROTATION
# ------------------------------------------------------------

def get_daily_topics(date_value: dt.date, count: int):
    seed = int(date_value.strftime("%Y%m%d"))
    rng = random.Random(seed)

    formulas = list(FORMULAS)
    rng.shuffle(formulas)

    return formulas[: min(count, len(formulas))]


# ------------------------------------------------------------
# EXCEL
# ------------------------------------------------------------

def excel_column_letter(number: int) -> str:
    result = ""

    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result

    return result


def write_cell(sheet, row, column, value):
    sheet.Cells(row, column).Value = value


def build_excel_workbook(topic, workbook_path: Path):
    log(f"Creating Excel workbook: {workbook_path.name}")

    excel = win32.DispatchEx("Excel.Application")

    excel.Visible = True
    excel.DisplayAlerts = False
    excel.ScreenUpdating = True

    workbook = excel.Workbooks.Add()

    # Keep one worksheet.
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()

    sheet = workbook.Worksheets(1)
    sheet.Name = "LearnVerse"

    headers = topic["headers"]

    # Title.
    sheet.Range("A1:G1").Merge()
    sheet.Range("A1").Value = f"Learn Verse Excel Tip — {topic['name']}"
    sheet.Range("A1").Font.Bold = True
    sheet.Range("A1").Font.Size = 18

    # Problem.
    sheet.Range("A2:G2").Merge()
    sheet.Range("A2").Value = topic["problem"]
    sheet.Range("A2").Font.Size = 11

    # Headers.
    for col, header in enumerate(headers, start=1):
        cell = sheet.Cells(4, col)
        cell.Value = header
        cell.Font.Bold = True

    # Data.
    for row_index, row_data in enumerate(topic["rows"], start=5):
        for col_index, value in enumerate(row_data, start=1):
            write_cell(sheet, row_index, col_index, value)

    # Formula demonstration row.
    demo_row = 12

    sheet.Range(f"A{demo_row}:F{demo_row}").Merge()
    sheet.Range(f"A{demo_row}").Value = "FORMULA DEMO"
    sheet.Range(f"A{demo_row}").Font.Bold = True
    sheet.Range(f"A{demo_row}").Font.Size = 14

    formula_cell = topic["formula_cell"]

    # Translate row-2 formulas to the actual demo row where possible.
    formula = topic["formula"]

    if "2" in formula:
        formula = formula.replace("2", str(demo_row + 1))

    if topic["id"] == "sum":
        formula = "=SUM(F5:F10)"

    elif topic["id"] == "average":
        formula = "=AVERAGE(F5:F10)"

    elif topic["id"] == "max":
        formula = "=MAX(F5:F10)"

    elif topic["id"] == "min":
        formula = "=MIN(F5:F10)"

    elif topic["id"] == "countif":
        formula = '=COUNTIF(C5:C10,"Done")'

    elif topic["id"] == "sumif":
        formula = '=SUMIF(C5:C10,"Done",F5:F10)'

    elif topic["id"] == "if":
        formula = '=IF(C5="Done","Completed","Pending")'

    elif topic["id"] == "left":
        formula = "=LEFT(F5,3)"

    elif topic["id"] == "round":
        formula = "=ROUND(F5,2)"

    elif topic["id"] == "product":
        formula = "=PRODUCT(F5,F6)"

    elif topic["id"] == "textjoin":
        formula = '=TEXTJOIN(" - ",TRUE,A5:C5)'

    elif topic["id"] == "xlookup":
        formula = '=XLOOKUP(A5,A5:A10,F5:F10,"Not Found")'

    # Use G13 for most final results so A:G remains visible.
    result_cell = sheet.Range("G13")
    result_cell.Formula = formula

    sheet.Range("F13").Value = "Formula:"
    sheet.Range("F13").Font.Bold = True

    sheet.Range("A15:G15").Merge()
    sheet.Range("A15").Value = topic["explanation"]
    sheet.Range("A15").WrapText = True

    # Formatting.
    sheet.Range("A4:G10").Borders.LineStyle = 1
    sheet.Range("A4:G4").Font.Bold = True

    for col in range(1, 8):
        sheet.Columns(col).ColumnWidth = 16

    sheet.Columns(1).ColumnWidth = 18
    sheet.Columns(2).ColumnWidth = 17
    sheet.Columns(3).ColumnWidth = 15
    sheet.Columns(4).ColumnWidth = 14
    sheet.Columns(5).ColumnWidth = 15
    sheet.Columns(6).ColumnWidth = 15
    sheet.Columns(7).ColumnWidth = 18

    sheet.Rows(1).RowHeight = 30
    sheet.Rows(2).RowHeight = 32

    # Freeze top rows.
    try:
        excel.ActiveWindow.SplitRow = 4
        excel.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    # Start at A1 and show A:G.
    try:
        excel.ActiveWindow.Zoom = EXCEL_ZOOM
        excel.ActiveWindow.ScrollColumn = 1
        excel.ActiveWindow.ScrollRow = 1
        sheet.Range("A1").Select()
    except Exception:
        pass

    workbook.SaveAs(str(workbook_path))

    return excel, workbook, sheet


def focus_excel(excel):
    try:
        excel.Visible = True
        excel.WindowState = -4137
    except Exception:
        pass

    try:
        excel.Activate()
    except Exception:
        pass

    time.sleep(1)


def restore_excel_view(excel, sheet):
    focus_excel(excel)

    try:
        sheet.Activate()
    except Exception:
        pass

    try:
        excel.ActiveWindow.Zoom = EXCEL_ZOOM
    except Exception:
        pass

    try:
        sheet.Range("A1").Select()
    except Exception:
        pass

    try:
        pyautogui.hotkey("ctrl", "home")
    except Exception:
        pass

    time.sleep(0.5)

    try:
        excel.ActiveWindow.ScrollColumn = 1
        excel.ActiveWindow.ScrollRow = 1
    except Exception:
        pass

    time.sleep(1)


# ------------------------------------------------------------
# MOUSE / TYPING DEMO
# ------------------------------------------------------------

def cell_click_and_type(excel, sheet, cell_address: str, text: str):
    focus_excel(excel)

    try:
        sheet.Range(cell_address).Select()
    except Exception:
        pass

    time.sleep(0.5)

    pyautogui.write(str(text), interval=0.04)
    pyautogui.press("enter")

    time.sleep(0.6)


def demonstrate_formula(excel, sheet, topic):
    log(f"Demonstrating {topic['name']}")

    # Show top-left of Excel.
    restore_excel_view(excel, sheet)

    # Click the formula result area.
    try:
        sheet.Range("G13").Select()
        time.sleep(0.6)
    except Exception:
        pass

    # Show formula bar activity by editing the actual formula.
    try:
        pyautogui.press("f2")
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")

        formula = topic["formula"]

        if topic["id"] == "sum":
            formula = "=SUM(F5:F10)"
        elif topic["id"] == "average":
            formula = "=AVERAGE(F5:F10)"
        elif topic["id"] == "max":
            formula = "=MAX(F5:F10)"
        elif topic["id"] == "min":
            formula = "=MIN(F5:F10)"
        elif topic["id"] == "countif":
            formula = '=COUNTIF(C5:C10,"Done")'
        elif topic["id"] == "sumif":
            formula = '=SUMIF(C5:C10,"Done",F5:F10)'
        elif topic["id"] == "if":
            formula = '=IF(C5="Done","Completed","Pending")'
        elif topic["id"] == "left":
            formula = "=LEFT(F5,3)"
        elif topic["id"] == "round":
            formula = "=ROUND(F5,2)"
        elif topic["id"] == "product":
            formula = "=PRODUCT(F5,F6)"
        elif topic["id"] == "textjoin":
            formula = '=TEXTJOIN(" - ",TRUE,A5:C5)'
        elif topic["id"] == "xlookup":
            formula = '=XLOOKUP(A5,A5:A10,F5:F10,"Not Found")'

        pyautogui.write(formula, interval=0.025)
        pyautogui.press("enter")

    except Exception as exc:
        log(f"Formula typing fallback: {exc}")

    time.sleep(1)

    restore_excel_view(excel, sheet)


# ------------------------------------------------------------
# VOICE
# ------------------------------------------------------------

def create_script(topic, index):
    return (
        f"Here's today's Excel tip. "
        f"Today we're learning {topic['name']}. "
        f"{topic['problem']} "
        f"First, prepare your Excel data. "
        f"Next, select the result cell. "
        f"Then enter {topic['formula']}. "
        f"Press Enter and Excel calculates the result. "
        f"{topic['explanation']} "
        f"Follow Learn Verse for more practical Excel tips."
    )


async def generate_voice_async(text: str, output: Path, voice: str):
    communicator = edge_tts.Communicate(
        text,
        voice=voice,
        rate=VOICE_RATE,
        volume=VOICE_VOLUME,
        pitch=VOICE_PITCH,
    )

    await communicator.save(str(output))


def generate_voice(text: str, output: Path, index: int):
    voice = VOICE_LIST[index % len(VOICE_LIST)]

    log(f"Generating voice: {voice}")

    asyncio.run(
        generate_voice_async(
            text,
            output,
            voice,
        )
    )

    if not output.exists() or output.stat().st_size < 1000:
        raise RuntimeError("Voice file was not generated correctly.")

    return voice


# ------------------------------------------------------------
# FFMPEG RECORDING
# ------------------------------------------------------------

def start_recording(output: Path, duration: float):
    require_program("ffmpeg")

    output = output.resolve()

    log(f"Starting Excel desktop recording for {duration:.2f} seconds.")

    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel",
        "warning",
        "-y",
        "-f",
        "gdigrab",
        "-framerate",
        str(FPS),
        "-draw_mouse",
        "1",
        "-video_size",
        f"{CAPTURE_WIDTH}x{CAPTURE_HEIGHT}",
        "-i",
        "desktop",
        "-t",
        f"{duration:.3f}",
        "-an",
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "18",
        "-pix_fmt",
        "yuv420p",
        str(output),
    ]

    log_file = output.with_name(output.stem + "_ffmpeg.log")

    handle = open(
        log_file,
        "w",
        encoding="utf-8",
    )

    process = subprocess.Popen(
        cmd,
        stdin=subprocess.DEVNULL,
        stdout=handle,
        stderr=subprocess.STDOUT,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
    )

    process._learnverse_log_handle = handle

    return process, log_file


def wait_recording(process, log_file: Path, output: Path):
    try:
        return_code = process.wait(timeout=MAX_VIDEO_SECONDS + 20)
    finally:
        try:
            handle = getattr(process, "_learnverse_log_handle", None)

            if handle:
                handle.close()
        except Exception:
            pass

    if return_code != 0:
        log_text = ""

        try:
            log_text = log_file.read_text(
                encoding="utf-8",
                errors="replace",
            )
        except Exception:
            pass

        raise RuntimeError(
            "FFmpeg recording failed.\n\n"
            f"FFmpeg log:\n{log_text}"
        )

    if not output.exists():
        raise RuntimeError("FFmpeg did not create the recording.")

    if output.stat().st_size < 50000:
        raise RuntimeError(
            f"Recording is too small: {output.stat().st_size} bytes"
        )

    log(
        f"Raw recording created: "
        f"{output.stat().st_size:,} bytes"
    )


# ------------------------------------------------------------
# FINAL VIDEO
# ------------------------------------------------------------

def escape_filter_text(text: str) -> str:
    return (
        str(text)
        .replace("\\", "\\\\")
        .replace(":", "\\:")
        .replace("'", "\\'")
        .replace(",", "\\,")
        .replace("[", "\\[")
        .replace("]", "\\]")
    )


def render_final_video(
    raw_video: Path,
    voice: Path,
    final_video: Path,
    topic,
):
    require_program("ffmpeg")
    require_program("ffprobe")

    voice_duration = ffprobe_duration(voice)

    if voice_duration <= 0:
        raise RuntimeError("Could not determine voice duration.")

    final_duration = min(
        MAX_VIDEO_SECONDS,
        max(1.0, voice_duration + 0.15),
    )

    log(
        f"Voice duration: {voice_duration:.2f}s | "
        f"Final duration: {final_duration:.2f}s"
    )

    title = escape_filter_text(topic["title"])
    formula_name = escape_filter_text(topic["name"])
    formula = escape_filter_text(topic["formula"])
    explanation = escape_filter_text(topic["explanation"])

    steps_text = "\\n".join(
        f"{i + 1}. {escape_filter_text(step)}"
        for i, step in enumerate(topic["steps"])
    )

    # Excel is captured in landscape.
    #
    # We keep the left 900 pixels, resize it into the upper
    # portion of a 1080x1920 vertical canvas, and use transparent
    # information cards below it.
    #
    # drawtext is intentionally used instead of external image
    # assets so the runner is self-contained.

    filter_complex = (
        "[0:v]"
        "crop=900:ih:0:0,"
        "scale=1080:1092:flags=lanczos,"
        "setsar=1,"
        "pad=1080:1920:0:0:white,"
        "format=yuv420p"
        "[base];"

        # Top title card.
        "[base]"
        f"drawbox=x=35:y=1120:w=1010:h=205:"
        "color=white@0.94:t=fill,"
        f"drawtext=text='{title}':"
        "fontcolor=black:"
        "fontsize=42:"
        "x=65:y=1145:"
        "borderw=1:"
        "bordercolor=white,"
        f"drawtext=text='Today's Steps':"
        "fontcolor=black:"
        "fontsize=30:"
        "x=65:y=1205,"
        f"drawtext=text='{steps_text}':"
        "fontcolor=black:"
        "fontsize=22:"
        "line_spacing=10:"
        "x=65:y=1245"
        "[card1];"

        # Formula card.
        "[card1]"
        "drawbox=x=35:y=1340:w=1010:h=245:"
        "color=white@0.94:t=fill,"
        f"drawtext=text='Formula Used':"
        "fontcolor=black:"
        "fontsize=30:"
        "x=65:y=1365,"
        f"drawtext=text='{formula_name}':"
        "fontcolor=black:"
        "fontsize=34:"
        "x=65:y=1410,"
        f"drawtext=text='{formula}':"
        "fontcolor=black:"
        "fontsize=24:"
        "x=65:y=1455,"
        f"drawtext=text='{explanation}':"
        "fontcolor=black:"
        "fontsize=20:"
        "line_spacing=8:"
        "x=65:y=1500"
        "[card2];"

        # YouTube branding card.
        "[card2]"
        "drawbox=x=35:y=1610:w=1010:h=220:"
        "color=white@0.94:t=fill,"
        "drawtext=text='Watch on YouTube':"
        "fontcolor=black:"
        "fontsize=32:"
        "x=65:y=1640,"
        "drawtext=text='LearnVerse9556':"
        "fontcolor=black:"
        "fontsize=38:"
        "x=65:y=1690,"
        "drawtext=text='Follow Learn Verse for more Excel tips':"
        "fontcolor=black:"
        "fontsize=24:"
        "x=65:y=1750"
        "[video];"

        # Audio cleanup.
        "[1:a]"
        "highpass=f=80,"
        "lowpass=f=15000,"
        "equalizer=f=1800:t=q:w=1:g=1.5,"
        "equalizer=f=3000:t=q:w=1:g=2,"
        "acompressor="
        "threshold=-19dB:"
        "ratio=2.5:"
        "attack=10:"
        "release=140:"
        "makeup=2,"
        "loudnorm=I=-15:TP=-1.5:LRA=6,"
        "aresample=48000"
        "[audio]"
    )

    cmd = [
        "ffmpeg",
        "-hide_banner",
        "-loglevel",
        "warning",
        "-y",
        "-i",
        str(raw_video),
        "-i",
        str(voice),
        "-filter_complex",
        filter_complex,
        "-map",
        "[video]",
        "-map",
        "[audio]",
        "-t",
        f"{final_duration:.3f}",
        "-c:v",
        "libx264",
        "-preset",
        "medium",
        "-crf",
        "19",
        "-pix_fmt",
        "yuv420p",
        "-c:a",
        "aac",
        "-b:a",
        "160k",
        "-movflags",
        "+faststart",
        str(final_video),
    ]

    run_command(cmd, timeout=180)

    if not final_video.exists():
        raise RuntimeError("Final video was not created.")

    if final_video.stat().st_size < 50000:
        raise RuntimeError(
            f"Final video is too small: "
            f"{final_video.stat().st_size} bytes"
        )

    actual_duration = ffprobe_duration(final_video)

    log(
        f"Final video: {final_video.name} | "
        f"{final_video.stat().st_size:,} bytes | "
        f"{actual_duration:.2f}s"
    )

    return actual_duration


# ------------------------------------------------------------
# METADATA
# ------------------------------------------------------------

def save_metadata(
    topic,
    index,
    date_value,
    final_video,
    workbook,
    voice_name,
    duration,
):
    metadata_path = OUTPUT / (
        f"learnverse_{date_value}_{index:02d}_{topic['id']}.json"
    )

    data = {
        "channel": "Learn Verse",
        "handle": "@LearnVerse9556",
        "date": str(date_value),
        "video_number": index,
        "formula_id": topic["id"],
        "formula_name": topic["name"],
        "title": topic["title"],
        "problem": topic["problem"],
        "formula": topic["formula"],
        "explanation": topic["explanation"],
        "steps": topic["steps"],
        "voice": voice_name,
        "duration_seconds": round(duration, 2),
        "video_file": final_video.name,
        "workbook_file": workbook.name,
        "hashtags": [
            "#Excel",
            "#ExcelTips",
            "#MicrosoftExcel",
            "#ExcelFormula",
            "#LearnVerse",
            "#Shorts",
        ],
        "youtube_title": (
            f"{topic['name']} Excel Formula "
            f"That Saves Time! #Excel #Shorts"
        ),
        "youtube_description": (
            f"Learn how to use Excel {topic['name']} "
            f"with a practical example.\n\n"
            f"Formula: {topic['formula']}\n\n"
            f"Follow Learn Verse for more practical Excel tips.\n"
            f"@LearnVerse9556"
        ),
    }

    metadata_path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    return metadata_path


# ------------------------------------------------------------
# SINGLE VIDEO
# ------------------------------------------------------------

def generate_one(topic, index, date_value):
    prefix = (
        f"learnverse_{date_value}_{index:02d}_{topic['id']}"
    )

    workbook_path = OUTPUT / f"{prefix}.xlsx"
    voice_path = OUTPUT / f"{prefix}.mp3"
    raw_path = OUTPUT / f"{prefix}_raw.mp4"
    final_path = OUTPUT / f"{prefix}.mp4"

    log("=" * 70)
    log(f"VIDEO {index}: {topic['title']}")
    log("=" * 70)

    excel = None
    workbook = None
    recorder = None

    try:
        # 1. Create workbook.
        excel, workbook, sheet = build_excel_workbook(
            topic,
            workbook_path,
        )

        # 2. Prepare Excel.
        focus_excel(excel)
        restore_excel_view(excel, sheet)

        # 3. Generate voice BEFORE recording so we know the duration.
        script = create_script(topic, index)

        voice_name = generate_voice(
            script,
            voice_path,
            index - 1,
        )

        voice_duration = ffprobe_duration(voice_path)

        if voice_duration <= 0:
            raise RuntimeError(
                "Voice duration could not be determined."
            )

        # 4. Recording duration follows narration.
        recording_duration = min(
            MAX_VIDEO_SECONDS,
            max(2.0, voice_duration + 0.8),
        )

        # 5. Start recording.
        recorder, log_file = start_recording(
            raw_path,
            recording_duration,
        )

        # Give FFmpeg a moment to start capturing.
        time.sleep(1.0)

        # 6. Perform the actual Excel interaction.
        demonstrate_formula(
            excel,
            sheet,
            topic,
        )

        # 7. Keep Excel visible at A:G for the final part.
        restore_excel_view(excel, sheet)

        # Wait for recording.
        wait_recording(
            recorder,
            log_file,
            raw_path,
        )

        recorder = None

        # 8. Render final vertical Short.
        duration = render_final_video(
            raw_path,
            voice_path,
            final_path,
            topic,
        )

        # 9. Save metadata.
        metadata_path = save_metadata(
            topic,
            index,
            date_value,
            final_path,
            workbook_path,
            voice_name,
            duration,
        )

        log(f"SUCCESS: {final_path.name}")

        return {
            "video": final_path,
            "workbook": workbook_path,
            "metadata": metadata_path,
        }

    finally:
        # Never leave FFmpeg running.
        if recorder is not None:
            try:
                if recorder.poll() is None:
                    recorder.terminate()
                    recorder.wait(timeout=10)
            except Exception:
                pass

        # Close Excel.
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        # Remove temporary raw recording and voice.
        safe_delete(raw_path)
        safe_delete(voice_path)


# ------------------------------------------------------------
# CLEANING
# ------------------------------------------------------------

def clean_temporary_files():
    for pattern in [
        "*_raw.mp4",
        "*_ffmpeg.log",
        "*.tmp",
    ]:
        for file in OUTPUT.glob(pattern):
            safe_delete(file)


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Learn Verse Excel Shorts Generator"
    )

    parser.add_argument(
        "--count",
        type=int,
        default=3,
        choices=range(1, 4),
        help="Number of Shorts to generate",
    )

    parser.add_argument(
        "--date",
        type=str,
        default=None,
        help="Optional date in YYYY-MM-DD format",
    )

    args = parser.parse_args()

    require_program("ffmpeg")
    require_program("ffprobe")

    if args.date:
        date_value = dt.datetime.strptime(
            args.date,
            "%Y-%m-%d",
        ).date()
    else:
        date_value = dt.date.today()

    log("=" * 70)
    log("LEARN VERSE EXCEL SHORTS")
    log("=" * 70)
    log(f"Date: {date_value}")
    log(f"Count: {args.count}")
    log(f"Output: {OUTPUT}")
    log("=" * 70)

    topics = get_daily_topics(
        date_value,
        args.count,
    )

    results = []

    try:
        for index, topic in enumerate(topics, start=1):
            try:
                result = generate_one(
                    topic,
                    index,
                    date_value,
                )

                results.append(result)

            except Exception as exc:
                log(
                    f"VIDEO {index} FAILED: "
                    f"{type(exc).__name__}: {exc}"
                )

                # Continue with the next video rather than
                # destroying the whole day's generation.
                continue

    finally:
        clean_temporary_files()

    log("=" * 70)
    log("GENERATION SUMMARY")
    log("=" * 70)

    for result in results:
        log(
            f"VIDEO: {result['video'].name}"
        )

        log(
            f"WORKBOOK: {result['workbook'].name}"
        )

        log(
            f"METADATA: {result['metadata'].name}"
        )

    log("=" * 70)
    log(
        f"Generated successfully: "
        f"{len(results)}/{args.count}"
    )
    log("=" * 70)

    if len(results) == 0:
        raise RuntimeError(
            "No Excel Shorts were generated successfully."
        )

    if len(results) < args.count:
        log(
            "WARNING: Some requested videos failed. "
            "Successful videos were preserved."
        )

    return 0


if __name__ == "__main__":
    sys.exit(main())