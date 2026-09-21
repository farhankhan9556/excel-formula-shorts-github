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
import win32con
import win32gui
import win32process
import win32api
import ctypes
import edge_tts
from PIL import Image, ImageDraw, ImageFont


# ============================================================
# LEARN VERSE - EXCEL SHORTS GENERATOR
# Design: Excel Quick Tip / branded educational layout v2
# Self-contained Windows + GitHub Actions version
# ============================================================

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
# When running inside GitHub Actions, this is the permanent PC folder used by
# the YouTube uploader. If it exists, copy final files there too.
PERMANENT_OUTPUT = Path(r"C:\excel-formula-shorts-github\output")
OUTPUT.mkdir(parents=True, exist_ok=True)

pyautogui.PAUSE = 0.15
pyautogui.FAILSAFE = False


# ------------------------------------------------------------
# SETTINGS
# ------------------------------------------------------------

# Edge-TTS can occasionally return NoAudioReceived for a particular voice.
# We therefore use adult female voices with automatic fallback and several
# safe speaking-rate attempts. The first successful voice/rate is kept.
VOICE_LIST = [
    "en-US-JennyNeural",
    "en-US-AriaNeural",
]
VOICE_RATES = ["+8%", "+12%", "+16%", "+20%"]
VOICE_VOLUME = "+0%"
VOICE_PITCH = "+1Hz"

# Keep the finished Short comfortably below 30 seconds without cutting the
# narration. If narration is still too long after the fastest safe rate, the
# generator fails that video rather than silently cutting spoken words.
TARGET_VOICE_SECONDS = 29.40
VOICE_RETRIES_PER_SETTING = 2

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
        stdin=subprocess.DEVNULL,
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
    # The formula/result is deliberately placed on row 11 so it remains
    # visible in the recorded Excel area after the six data rows.
    formula_cell = topic["formula_cell"]
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

    # Put the live formula/result immediately below the six data rows.
    # This keeps the formula cell visible inside the recorded Excel area.
    formula_row = 11
    result_cell = sheet.Range(f"G{formula_row}")
    result_cell.Formula = formula

    sheet.Range(f"F{formula_row}").Value = "Formula:"
    sheet.Range(f"F{formula_row}").Font.Bold = True

    # Compact explanation row so more of the worksheet remains visible.
    sheet.Range("A12:G12").Merge()
    sheet.Range("A12").Value = topic["explanation"]
    sheet.Range("A12").WrapText = True

    # --------------------------------------------------------
    # Professional Excel formatting
    # --------------------------------------------------------
    # Keep the complete working area A:G visible and hide all
    # columns to the right so the recording cannot drift into H+.
    try:
        sheet.Columns("H:XFD").Hidden = True
    except Exception:
        pass

    # Thin, professional table borders.
    table_range = sheet.Range("A4:G10")
    table_range.Borders.LineStyle = 1
    table_range.Borders.Weight = 2

    # Center all table content horizontally and vertically.
    table_range.HorizontalAlignment = -4108  # xlCenter
    table_range.VerticalAlignment = -4108    # xlCenter

    # Header styling.
    header_range = sheet.Range("A4:G4")
    header_range.Font.Bold = True
    header_range.Font.Size = 10
    header_range.HorizontalAlignment = -4108
    header_range.VerticalAlignment = -4108

    # Compact readable fonts.
    sheet.Range("A1:G1").Font.Size = 14
    sheet.Range("A2:G2").Font.Size = 9
    sheet.Range("A5:G10").Font.Size = 10
    sheet.Range("F11:G11").Font.Size = 9
    sheet.Range("A12:G12").Font.Size = 8

    # Professional blue/green table palette.
    header_range.Interior.Color = 0xD07A08
    header_range.Font.Color = 0xFFFFFF

    # Alternating data rows.
    for r in range(5, 11):
        row_range = sheet.Range(f"A{r}:G{r}")
        if r % 2 == 1:
            row_range.Interior.Color = 0xF2F7FC
        else:
            row_range.Interior.Color = 0xFFFFFF

    # Color-code Status (column C) and Priority (column D).
    for r in range(5, 11):
        status = str(sheet.Cells(r, 3).Value or "").strip().lower()
        priority = str(sheet.Cells(r, 4).Value or "").strip().lower()

        if status == "done":
            sheet.Cells(r, 3).Interior.Color = 0xD9F2E3
            sheet.Cells(r, 3).Font.Color = 0x246B3A
            sheet.Cells(r, 3).Font.Bold = True
        elif status == "pending":
            sheet.Cells(r, 3).Interior.Color = 0xD9EAF7
            sheet.Cells(r, 3).Font.Color = 0x8A5A00
            sheet.Cells(r, 3).Font.Bold = True

        if priority == "high":
            sheet.Cells(r, 4).Interior.Color = 0xD9D9FF
            sheet.Cells(r, 4).Font.Bold = True
        elif priority == "medium":
            sheet.Cells(r, 4).Interior.Color = 0xFFF0C2
            sheet.Cells(r, 4).Font.Bold = True
        elif priority == "low":
            sheet.Cells(r, 4).Interior.Color = 0xDCEBFF
            sheet.Cells(r, 4).Font.Bold = True

    # Formula row: clearly visible and centered.
    sheet.Range("F11:G11").Interior.Color = 0xD9F2E3
    sheet.Range("F11:G11").Borders.LineStyle = 1
    sheet.Range("F11:G11").Borders.Weight = 2
    sheet.Range("F11:G11").HorizontalAlignment = -4108
    sheet.Range("F11:G11").VerticalAlignment = -4108
    sheet.Range("G11").Font.Bold = True

    # Explanation row.
    sheet.Range("A12:G12").Interior.Color = 0xE8F5FF
    sheet.Range("A12:G12").Font.Color = 0x203040
    sheet.Range("A12:G12").HorizontalAlignment = -4108
    sheet.Range("A12:G12").VerticalAlignment = -4108

    # Column widths deliberately compact so A:G fit in the captured area.
    widths = {
        1: 15, 2: 14, 3: 13, 4: 12, 5: 12, 6: 13, 7: 15
    }
    for col, width in widths.items():
        sheet.Columns(col).ColumnWidth = width

    # Compact first rows so the data and live formula row appear together.
    sheet.Rows(1).RowHeight = 22
    sheet.Rows(2).RowHeight = 20
    sheet.Rows(3).RowHeight = 5
    sheet.Rows(4).RowHeight = 21
    for r in range(5, 11):
        sheet.Rows(r).RowHeight = 21
    sheet.Rows(11).RowHeight = 22
    sheet.Rows(12).RowHeight = 24

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


def _force_window_foreground(hwnd):
    """Force Excel to the primary monitor, maximized and foreground/topmost."""
    if not hwnd:
        return False

    try:
        # Restore first so SetWindowPos can establish a deterministic size.
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.15)

        # Get the primary monitor work area.
        monitor = win32api.GetMonitorInfo(
            win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTOPRIMARY)
        )
        left, top, right, bottom = monitor["Work"]
        width = right - left
        height = bottom - top

        # Temporarily make Excel topmost AND explicitly fill the work area.
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOPMOST,
            left,
            top,
            width,
            height,
            win32con.SWP_SHOWWINDOW,
        )

        # Ask Windows to maximize as well.
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)

        # Foreground handling: Windows may otherwise reject SetForegroundWindow
        # when the runner's process is not the foreground owner.
        foreground = win32gui.GetForegroundWindow()
        current_tid = (
            win32process.GetWindowThreadProcessId(foreground)[0]
            if foreground else 0
        )
        target_tid = win32process.GetWindowThreadProcessId(hwnd)[0]

        attached = False
        if current_tid and current_tid != target_tid:
            try:
                win32process.AttachThreadInput(current_tid, target_tid, True)
                attached = True
            except Exception:
                pass

        try:
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
            try:
                ctypes.windll.user32.SetActiveWindow(hwnd)
            except Exception:
                pass
        finally:
            if attached:
                try:
                    win32process.AttachThreadInput(current_tid, target_tid, False)
                except Exception:
                    pass

        # Final deterministic bounds check after maximizing.
        try:
            win32gui.SetWindowPos(
                hwnd,
                win32con.HWND_TOPMOST,
                left,
                top,
                width,
                height,
                win32con.SWP_SHOWWINDOW,
            )
        except Exception:
            pass

        time.sleep(0.25)
        return win32gui.GetForegroundWindow() == hwnd

    except Exception as exc:
        log(f"Excel foreground enforcement warning: {exc}")
        return False


def _excel_hwnd(excel):
    try:
        return int(excel.Hwnd)
    except Exception:
        return 0


def focus_excel(excel):
    try:
        excel.Visible = True
    except Exception:
        pass

    hwnd = _excel_hwnd(excel)

    if hwnd:
        for _ in range(3):
            if _force_window_foreground(hwnd):
                break
            time.sleep(0.25)

    try:
        excel.WindowState = -4137  # xlMaximized
    except Exception:
        pass

    try:
        excel.Activate()
    except Exception:
        pass

    time.sleep(0.5)

    if hwnd:
        _force_window_foreground(hwnd)

    time.sleep(0.8)


def keep_excel_on_top(excel):
    """Re-assert Excel as the maximized foreground/topmost window."""
    try:
        excel.Visible = True
    except Exception:
        pass

    hwnd = _excel_hwnd(excel)
    if not hwnd:
        return

    _force_window_foreground(hwnd)

    # GUI fallback: Alt+Space -> X is the standard Windows maximize command.
    # This is deliberately done only while Excel is already foreground.
    try:
        if win32gui.GetForegroundWindow() == hwnd:
            pyautogui.hotkey("alt", "space")
            time.sleep(0.2)
            pyautogui.press("x")
            time.sleep(0.5)
    except Exception as exc:
        log(f"Excel GUI maximize fallback warning: {exc}")

    # Re-assert topmost after the GUI maximize command.
    _force_window_foreground(hwnd)


def restore_excel_view(excel, sheet):
    focus_excel(excel)

    try:
        sheet.Activate()
    except Exception:
        pass

    try:
        excel.ActiveWindow.WindowState = -4137  # xlMaximized
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

    time.sleep(0.4)

    try:
        excel.ActiveWindow.ScrollColumn = 1
        excel.ActiveWindow.ScrollRow = 1
    except Exception:
        pass

    try:
        sheet.Range("A1:G1").Select()
    except Exception:
        pass

    keep_excel_on_top(excel)
    time.sleep(0.5)


def verify_excel_recording_window(excel):
    """Verify Excel is foreground and fills the primary monitor before capture."""
    hwnd = _excel_hwnd(excel)
    if not hwnd:
        raise RuntimeError("Could not obtain Excel window handle.")

    keep_excel_on_top(excel)

    foreground = win32gui.GetForegroundWindow()
    if foreground != hwnd:
        raise RuntimeError(
            f"Excel is not foreground before recording. "
            f"Excel HWND={hwnd}, foreground HWND={foreground}"
        )

    rect = win32gui.GetWindowRect(hwnd)
    monitor = win32api.GetMonitorInfo(
        win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTOPRIMARY)
    )
    left, top, right, bottom = monitor["Work"]
    win_w = rect[2] - rect[0]
    win_h = rect[3] - rect[1]
    mon_w = right - left
    mon_h = bottom - top

    # Allow a small Windows border tolerance.
    if win_w < mon_w - 20 or win_h < mon_h - 20:
        raise RuntimeError(
            f"Excel is not maximized. Window={win_w}x{win_h}, "
            f"monitor work area={mon_w}x{mon_h}"
        )

    log(
        f"Excel recording window verified: foreground=True, "
        f"window={win_w}x{win_h}, monitor={mon_w}x{mon_h}"
    )



def release_excel_topmost(excel):
    """Return Excel to normal z-order before closing it."""
    hwnd = _excel_hwnd(excel)
    if not hwnd:
        return

    try:
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_NOTOPMOST,
            0,
            0,
            0,
            0,
            win32con.SWP_NOMOVE
            | win32con.SWP_NOSIZE
            | win32con.SWP_SHOWWINDOW,
        )
    except Exception:
        pass

def demonstrate_formula(excel, sheet, topic):
    log(f"Demonstrating {topic['name']}")

    # Show top-left of Excel.
    restore_excel_view(excel, sheet)

    # Re-assert Excel immediately before any mouse/keyboard automation.
    # This prevents another application/dialog from receiving keystrokes.
    keep_excel_on_top(excel)

    # Click the formula result area.
    try:
        sheet.Range("G11").Select()
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

    keep_excel_on_top(excel)
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


async def generate_voice_async(text: str, output: Path, voice: str, rate: str):
    communicator = edge_tts.Communicate(
        text,
        voice=voice,
        rate=rate,
        volume=VOICE_VOLUME,
        pitch=VOICE_PITCH,
    )

    await communicator.save(str(output))


def generate_voice(text: str, output: Path, index: int):
    # Rotate the preferred voice by video number, but automatically fall back
    # to the other adult female voices when a provider request fails. This
    # prevents one temporary Edge-TTS voice error from losing an entire day.
    preferred = VOICE_LIST[index % len(VOICE_LIST)]
    ordered_voices = [preferred] + [v for v in VOICE_LIST if v != preferred]

    last_error = None

    for voice in ordered_voices:
        for rate in VOICE_RATES:
            for attempt in range(1, VOICE_RETRIES_PER_SETTING + 1):
                safe_delete(output)
                log(
                    f"Generating voice: {voice} | rate {rate} | "
                    f"attempt {attempt}/{VOICE_RETRIES_PER_SETTING}"
                )

                try:
                    asyncio.run(
                        generate_voice_async(
                            text,
                            output,
                            voice,
                            rate,
                        )
                    )

                    if not output.exists() or output.stat().st_size < 1000:
                        raise RuntimeError("Voice file was not generated correctly.")

                    duration = ffprobe_duration(output)
                    if duration <= 0:
                        raise RuntimeError("Voice duration could not be determined.")

                    if duration <= TARGET_VOICE_SECONDS:
                        log(
                            f"Voice ready: {voice} | {rate} | "
                            f"{duration:.2f}s"
                        )
                        return voice

                    log(
                        f"Voice is {duration:.2f}s; trying a faster rate "
                        f"without cutting narration."
                    )

                except Exception as exc:
                    last_error = exc
                    log(
                        f"Voice attempt failed: {voice} / {rate}: "
                        f"{type(exc).__name__}: {exc}"
                    )
                    time.sleep(0.8)

    safe_delete(output)
    raise RuntimeError(
        "All configured Edge-TTS voices/rates failed or exceeded the "
        f"{TARGET_VOICE_SECONDS:.2f}s narration target. "
        f"Last error: {last_error}"
    )


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


def _font(size, bold=False):
    candidates = []
    if bold:
        candidates += [
            r"C:\\Windows\\Fonts\\arialbd.ttf",
            r"C:\\Windows\\Fonts\\segoeuib.ttf",
        ]
    candidates += [
        r"C:\\Windows\\Fonts\\arial.ttf",
        r"C:\\Windows\\Fonts\\segoeui.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                pass
    return ImageFont.load_default()


def _wrap_text(draw, text, font, max_width):
    words = str(text).split()
    lines = []
    current = ""
    for word in words:
        trial = word if not current else current + " " + word
        box = draw.textbbox((0, 0), trial, font=font)
        if box[2] - box[0] <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines



def _hex(hex_value):
    hex_value = hex_value.lstrip("#")
    return tuple(int(hex_value[i:i + 2], 16) for i in (0, 2, 4)) + (255,)


def _fit_font(draw, text, font_path_candidates, max_size, min_size, max_width, bold=False):
    for size in range(max_size, min_size - 1, -2):
        font = _font(size, bold)
        box = draw.textbbox((0, 0), str(text), font=font)
        if box[2] - box[0] <= max_width:
            return font
    return _font(min_size, bold)


def _draw_icon_circle(draw, center, radius, fill, text_value, font):
    x, y = center
    draw.ellipse(
        (x - radius, y - radius, x + radius, y + radius),
        fill=fill,
    )
    box = draw.textbbox((0, 0), text_value, font=font)
    tw = box[2] - box[0]
    th = box[3] - box[1]
    draw.text(
        (x - tw / 2, y - th / 2 - 2),
        text_value,
        font=font,
        fill=(255, 255, 255, 255),
    )


def _draw_youtube_icon(draw, x, y, w=72, h=48):
    red = _hex("#FF1F2D")
    draw.rounded_rectangle((x, y, x + w, y + h), radius=12, fill=red)
    draw.polygon(
        [(x + 29, y + 12), (x + 29, y + 36), (x + 49, y + 24)],
        fill=(255, 255, 255, 255),
    )


def _result_for_topic(topic):
    tid = topic["id"]
    rows = topic.get("rows", [])
    values_f = [r[5] for r in rows if len(r) > 5 and isinstance(r[5], (int, float))]
    if tid == "if":
        return "Completed", "Status is Done → Completed"
    if tid == "sum":
        return f"{sum(values_f):,.0f}", "Total of the selected sales values"
    if tid == "average":
        return f"{sum(values_f) / len(values_f):,.2f}", "Average of the selected values"
    if tid == "max":
        return f"{max(values_f):,.0f}", "Highest value in the selected range"
    if tid == "min":
        return f"{min(values_f):,.0f}", "Lowest value in the selected range"
    if tid == "countif":
        return str(sum(1 for r in rows if len(r) > 2 and r[2] == "Done")), "Rows where Status = Done"
    if tid == "sumif":
        return f"{sum(r[5] for r in rows if len(r) > 5 and r[2] == 'Done'):,.0f}", "Sales from rows where Status = Done"
    if tid == "round":
        return f"{float(rows[0][5]):,.2f}", "Rounded to 2 decimal places"
    if tid == "left":
        return str(rows[0][5])[:3], "First 3 characters from the code"
    if tid == "product":
        return f"{rows[0][5] * rows[1][5]:,.0f}", "Quantity × Price"
    if tid == "textjoin":
        return f"{rows[0][0]} - {rows[0][1]} - {rows[0][2]}", "Combined text with a separator"
    if tid == "xlookup":
        return str(rows[0][5]), "Matching value returned from the lookup"
    return "Done", "Formula result shown in Excel"


def _pro_tips(topic):
    tips = {
        "if": ["Use IF for status checks.", "Combine IF with AND/OR.", "Keep result text short."],
        "sum": ["Works across rows or columns.", "Use SUM for totals.", "Avoid manual addition."],
        "average": ["Works with numbers and scores.", "Use a clean range.", "Combine with IF for analysis."],
        "max": ["Useful for highest sales.", "Works with dates too.", "Great for quick comparisons."],
        "min": ["Useful for lowest values.", "Works with dates too.", "Great for exception checks."],
        "countif": ["Perfect for status counts.", "Use wildcards for text.", "Great for dashboards."],
        "sumif": ["Add only matching values.", "Use a clear criteria range.", "Great for sales reports."],
        "round": ["Control decimal places.", "Useful for prices.", "Avoid long decimal displays."],
        "left": ["Useful for codes.", "Change 3 to any count.", "Great for text cleanup."],
        "product": ["Multiply values quickly.", "Useful for quantity × price.", "Works with cell references."],
        "textjoin": ["Add any delimiter.", "TRUE ignores empty cells.", "Great for labels and IDs."],
        "xlookup": ["Use exact matches.", "Can return text or numbers.", "Great for lookup tables."],
    }
    return tips.get(topic["id"], ["Use a clean data range.", "Check your references.", "Practice with real data."])


def create_overlay_png(topic, path: Path):
    """Create a polished Learn Verse educational overlay matching the supplied
    Excel Quick Tip reference: branded header, topic banner, three information
    cards, step-by-step panel, pro tips and YouTube footer.
    """
    W, H = 1080, 1920
    image = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)

    navy = _hex("#073B78")
    blue = _hex("#0878D1")
    light_blue = _hex("#E8F5FF")
    green = _hex("#079447")
    light_green = _hex("#E9F9EF")
    yellow = _hex("#FFC928")
    light_yellow = _hex("#FFF8DD")
    purple = _hex("#7456D8")
    pink = _hex("#E94E77")
    red = _hex("#FF1F2D")
    dark = _hex("#102A43")
    gray = _hex("#52677D")
    white = (255, 255, 255, 255)

    # ---------- Header ----------
    draw.rectangle((0, 0, W, 150), fill=(246, 251, 255, 255))

    # Excel-style logo.
    draw.rounded_rectangle((28, 30, 132, 130), radius=18, fill=_hex("#107C41"))
    draw.rounded_rectangle((48, 48, 112, 112), radius=10, fill=_hex("#21A366"))
    xfont = _font(54, True)
    draw.text((59, 49), "X", font=xfont, fill=white)

    header_font = _fit_font(draw, "Excel Quick Tip", [], 58, 40, 610, True)
    draw.text((165, 23), "Excel Quick Tip", font=header_font, fill=navy)

    sub_font = _font(31, False)
    draw.text((168, 86), "Simple Formulas  •  Big Results", font=sub_font, fill=navy)
    draw.line((168, 124, 545, 124), fill=yellow, width=5)

    # Lightbulb icon.
    draw.ellipse((605, 36, 661, 92), outline=yellow, width=6)
    draw.rectangle((620, 88, 646, 103), fill=yellow)
    draw.line((615, 110, 651, 110), fill=navy, width=4)

    # Learn Verse branding.
    brand_big = _font(32, True)
    brand_small = _font(18, False)
    draw.text((720, 28), "LEARN VERSE", font=brand_big, fill=navy)
    draw.text((724, 72), "Learn • Practice • Grow", font=brand_small, fill=gray)

    # ---------- Topic banner ----------
    draw.rounded_rectangle((18, 162, W - 18, 300), radius=28, fill=green)
    num_font = _font(30, True)
    draw.rounded_rectangle((34, 184, 118, 270), radius=18, fill=_hex("#056B35"))
    draw.text((58, 207), "#1", font=num_font, fill=white)

    topic_title = f"{topic['name']} Formula in Excel"
    title_font = _fit_font(draw, topic_title, [], 48, 30, 675, True)
    draw.text((138, 180), topic_title, font=title_font, fill=yellow)

    problem_font = _fit_font(draw, topic["problem"], [], 27, 19, 690, False)
    draw.text((140, 235), topic["problem"], font=problem_font, fill=white)

    save_font = _font(20, False)
    draw.text((812, 188), "Save Time", font=save_font, fill=white)
    draw.text((812, 216), "Work Smarter", font=save_font, fill=white)
    draw.text((812, 244), "Be Productive", font=save_font, fill=white)

    # ---------- Excel reveal frame ----------
    draw.rounded_rectangle((18, 315, W - 18, 1040), radius=20, fill=(255, 255, 255, 0), outline=_hex("#D5DEE8"), width=3)
    # This border/frame is underneath the actual Excel capture.

    # ---------- Three detail cards ----------
    card_y1, card_y2 = 1060, 1250
    gap = 14
    left = 18
    card_w = (W - 36 - gap * 2) // 3
    xs = [left, left + card_w + gap, left + (card_w + gap) * 2]

    # Formula card
    draw.rounded_rectangle((xs[0], card_y1, xs[0] + card_w, card_y2), radius=22, fill=blue)
    icon_font = _font(25, True)
    _draw_icon_circle(draw, (xs[0] + 48, card_y1 + 48), 30, _hex("#0B5CA8"), "ƒ", icon_font)
    draw.text((xs[0] + 88, card_y1 + 25), "Formula Used", font=_font(25, True), fill=white)
    formula_font = _fit_font(draw, topic["formula"], [], 25, 15, card_w - 30, True)
    draw.rounded_rectangle((xs[0] + 18, card_y1 + 78, xs[0] + card_w - 18, card_y1 + 132),
                           radius=12, fill=white)
    draw.text((xs[0] + 30, card_y1 + 91), topic["formula"], font=formula_font, fill=navy)

    # What it does card
    draw.rounded_rectangle((xs[1], card_y1, xs[1] + card_w, card_y2), radius=22, fill=light_green)
    _draw_icon_circle(draw, (xs[1] + 48, card_y1 + 48), 30, green, "i", icon_font)
    draw.text((xs[1] + 88, card_y1 + 25), "What It Does", font=_font(25, True), fill=dark)
    desc = _wrap_text(draw, topic["explanation"], _font(19, False), card_w - 36)
    yy = card_y1 + 84
    for line in desc[:4]:
        draw.text((xs[1] + 18, yy), line, font=_font(19, False), fill=gray)
        yy += 27

    # Result card
    result, result_desc = _result_for_topic(topic)
    draw.rounded_rectangle((xs[2], card_y1, xs[2] + card_w, card_y2), radius=22, fill=_hex("#F0ECFF"))
    _draw_icon_circle(draw, (xs[2] + 48, card_y1 + 48), 30, purple, "✓", icon_font)
    draw.text((xs[2] + 88, card_y1 + 25), "Result", font=_font(25, True), fill=dark)
    result_font = _fit_font(draw, result, [], 34, 22, card_w - 36, True)
    draw.rounded_rectangle((xs[2] + 18, card_y1 + 75, xs[2] + card_w - 18, card_y1 + 128),
                           radius=14, fill=_hex("#DDF8E7"), outline=_hex("#56C982"), width=2)
    rbox = draw.textbbox((0, 0), result, font=result_font)
    rw = rbox[2] - rbox[0]
    draw.text((xs[2] + (card_w - rw) / 2, card_y1 + 87), result, font=result_font, fill=green)
    rd = _wrap_text(draw, result_desc, _font(16, False), card_w - 36)
    yy = card_y1 + 142
    for line in rd[:2]:
        draw.text((xs[2] + 18, yy), line, font=_font(16, False), fill=gray)
        yy += 22

    # ---------- Step by Step ----------
    sy1, sy2 = 1270, 1570
    draw.rounded_rectangle((18, sy1, W - 18, sy2), radius=24, fill=light_yellow, outline=yellow, width=3)
    _draw_icon_circle(draw, (68, sy1 + 52), 30, navy, "✓", icon_font)
    draw.text((116, sy1 + 24), "Step by Step", font=_font(34, True), fill=navy)
    draw.line((116, sy1 + 69, 350, sy1 + 69), fill=blue, width=4)

    steps = topic["steps"][:4]
    step_y = sy1 + 92
    step_colors = [red, blue, purple, green]
    for idx, step in enumerate(steps, 1):
        _draw_icon_circle(draw, (65, step_y + 17), 19, step_colors[idx - 1], str(idx), _font(18, True))
        step_font = _fit_font(draw, step, [], 23, 17, 860, False)
        draw.text((102, step_y + 3), step, font=step_font, fill=dark)
        step_y += 50

    # Highlight formula action if there is room.
    formula_label = "Formula: " + topic["formula"]
    ff = _fit_font(draw, formula_label, [], 20, 14, 820, True)
    draw.rounded_rectangle((102, sy2 - 52, 900, sy2 - 18), radius=10, fill=_hex("#D8F4E2"))
    draw.text((116, sy2 - 47), formula_label, font=ff, fill=green)

    # ---------- Pro Tips ----------
    py1, py2 = 1590, 1720
    draw.rounded_rectangle((18, py1, W - 18, py2), radius=22, fill=light_blue, outline=_hex("#7CC7FF"), width=3)
    _draw_icon_circle(draw, (68, py1 + 43), 28, navy, "★", _font(18, True))
    draw.text((110, py1 + 20), "Pro Tips", font=_font(30, True), fill=navy)

    tips = _pro_tips(topic)
    col_w = 290
    for i, tip in enumerate(tips):
        x = 120 + i * 315
        draw.ellipse((x, py1 + 62, x + 22, py1 + 84), fill=blue)
        draw.text((x + 5, py1 + 61), "✓", font=_font(14, True), fill=white)
        lines = _wrap_text(draw, tip, _font(17, False), col_w)
        yy = py1 + 58
        for line in lines[:3]:
            draw.text((x + 32, yy), line, font=_font(17, False), fill=dark)
            yy += 22

    # ---------- YouTube footer ----------
    fy1, fy2 = 1740, 1920
    draw.rounded_rectangle((0, fy1, W, fy2), radius=0, fill=navy)
    _draw_youtube_icon(draw, 45, fy1 + 28, 70, 48)
    draw.text((135, fy1 + 24), "Watch on YouTube", font=_font(34, True), fill=white)
    draw.text((135, fy1 + 73), "Learn more Excel tips, formulas and shortcuts", font=_font(18, False), fill=white)
    draw.text((135, fy1 + 101), "on our YouTube channel.", font=_font(18, False), fill=white)

    # Channel pill
    draw.rounded_rectangle((675, fy1 + 28, 1035, fy1 + 82), radius=27, fill=red)
    _draw_youtube_icon(draw, 690, fy1 + 35, 44, 34)
    draw.text((748, fy1 + 39), "LearnVerse9556", font=_font(24, True), fill=white)
    draw.text((748, fy1 + 91), "♧  Subscribe for more!", font=_font(18, False), fill=white)

    draw.line((135, fy1 + 132, 500, fy1 + 132), fill=blue, width=3)
    draw.text((310, fy1 + 142), "— Learn Verse • Excel Made Easy —", font=_font(19, False), fill=white)

    image.save(path, "PNG")



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

    overlay_path = final_video.with_name(final_video.stem + "_overlay.png")
    create_overlay_png(topic, overlay_path)

    # Landscape Excel capture -> vertical 1080x1920 canvas.
    # The lower cards are supplied as a transparent PNG, which avoids
    # FFmpeg drawtext quoting/filter-parser problems.
    # The reference design uses a branded header, a topic banner, a large
    # real-Excel area, then educational cards below it. The Excel capture is
    # placed between y=315 and y=1040; the transparent overlay supplies the
    # header/cards/footer.
    video_filter = (
        "[0:v]"
        "crop=900:640:0:0,"
        "scale=1080:770:flags=lanczos,"
        "setsar=1,"
        "pad=1080:1920:0:295:color=white,"
        "format=yuv420p"
        "[base];"
        "[base][2:v]overlay=0:0:format=auto,"
        "format=yuv420p"
        "[video]"
    )

    audio_filter = (
        "[1:a]"
        "highpass=f=80,"
        # Edge-TTS MP3 commonly uses a 24 kHz sample rate, so 10 kHz keeps
        # the low-pass frequency safely below Nyquist and avoids filter errors.
        "lowpass=f=10000,"
        "equalizer=f=1800:t=q:w=1:g=1.5,"
        "equalizer=f=3000:t=q:w=1:g=2,"
        "acompressor=threshold=-19dB:ratio=2.5:attack=10:release=140:makeup=2,"
        "loudnorm=I=-15:TP=-1.5:LRA=6,"
        "aresample=48000"
        "[audio]"
    )

    filter_complex = video_filter + ";" + audio_filter

    cmd = [
        "ffmpeg",
        "-nostdin",
        "-hide_banner",
        "-loglevel", "warning",
        "-y",
        "-i", str(raw_video),
        "-i", str(voice),
        "-loop", "1",
        "-framerate", str(FPS),
        "-i", str(overlay_path),
        "-filter_complex", filter_complex,
        "-map", "[video]",
        "-map", "[audio]",
        "-t", f"{final_duration:.3f}",
        "-c:v", "libx264",
        "-preset", "medium",
        "-crf", "19",
        "-pix_fmt", "yuv420p",
        "-c:a", "aac",
        "-b:a", "160k",
        "-movflags", "+faststart",
        str(final_video),
    ]

    try:
        run_command(cmd, timeout=180)
    finally:
        safe_delete(overlay_path)

    if not final_video.exists():
        raise RuntimeError("Final video was not created.")

    if final_video.stat().st_size < 50000:
        raise RuntimeError(
            f"Final video is too small: {final_video.stat().st_size} bytes"
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
    log(f"Generator output directory: {OUTPUT.resolve()}")
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

        # FFmpeg captures the whole desktop, so Excel must be the actual
        # foreground/maximized window before the demonstration begins.
        keep_excel_on_top(excel)
        time.sleep(0.8)
        verify_excel_recording_window(excel)

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

        if not final_path.exists() or final_path.stat().st_size < 50000:
            raise RuntimeError(
                f"Final video was not created correctly: {final_path}"
            )

        # Explicitly persist generated files to the permanent PC output folder.
        # This makes the generator independent of GitHub Actions' checkout path.
        if str(PERMANENT_OUTPUT.resolve()).lower() != str(OUTPUT.resolve()).lower():
            try:
                PERMANENT_OUTPUT.mkdir(parents=True, exist_ok=True)
                import shutil
                for persist_file in (
                    final_path,
                    workbook_path,
                    metadata_path,
                ):
                    if persist_file.exists():
                        destination = PERMANENT_OUTPUT / persist_file.name
                        shutil.copy2(persist_file, destination)
                        if not destination.exists():
                            raise RuntimeError(
                                f"Permanent copy missing after copy: {destination}"
                            )
                        log(f"PERMANENT COPY: {destination}")
            except Exception as exc:
                raise RuntimeError(
                    f"Could not persist generated files to "
                    f"{PERMANENT_OUTPUT}: {exc}"
                )

        log(
            f"SUCCESS: {final_path.name} | "
            f"{final_path.stat().st_size:,} bytes | "
            f"workspace={final_path.resolve()} | "
            f"permanent={PERMANENT_OUTPUT.resolve()}"
        )

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

        # Return Excel to normal z-order before closing.
        if excel is not None:
            release_excel_topmost(excel)

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
        "*_overlay.png",
        "*.tmp",
    ]:
        for file in OUTPUT.glob(pattern):
            safe_delete(file)


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Learn Verse Excel Shorts Generator")
    parser.add_argument("--count", type=int, default=3, choices=range(1, 4))
    parser.add_argument("--date", type=str, default=None)
    args = parser.parse_args()

    require_program("ffmpeg")
    require_program("ffprobe")

    if args.date:
        date_value = dt.datetime.strptime(args.date, "%Y-%m-%d").date()
    else:
        date_value = dt.date.today()

    log("=" * 70)
    log("LEARN VERSE EXCEL SHORTS")
    log("=" * 70)
    log(f"Date: {date_value}")
    log(f"Count: {args.count}")
    log(f"Output: {OUTPUT}")
    log("=" * 70)

    topics = get_daily_topics(date_value, args.count)
    results = []

    for index, topic in enumerate(topics, start=1):
        try:
            results.append(generate_one(topic, index, date_value))
        except Exception as exc:
            log(f"VIDEO {index} FAILED: {type(exc).__name__}: {exc}")
            continue

    clean_temporary_files()

    log("=" * 70)
    log(f"Generated successfully: {len(results)}/{args.count}")
    log("=" * 70)

    if results:
        for result in results:
            log(f"VIDEO: {result['video'].name}")
            log(f"WORKBOOK: {result['workbook'].name}")
            log(f"METADATA: {result['metadata'].name}")

    if not results:
        log("No video was generated. Review the VIDEO FAILED message above.")
        return 1

    if len(results) < args.count:
        log("WARNING: Some requested videos failed; successful videos were preserved.")

    return 0


if __name__ == "__main__":
    sys.exit(main())