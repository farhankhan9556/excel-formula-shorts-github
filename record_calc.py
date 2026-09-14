#!/usr/bin/env python3
"""
Open the genuine XLSX in LibreOffice Calc and record visible interaction.

Environment:
  DISPLAY=:99
  RECORD_SECONDS=8

The screen is captured by FFmpeg from the Xvfb display.
"""

import os
import subprocess
import sys
import time
from pathlib import Path

xlsx = Path(sys.argv[1]).resolve()
video = Path(sys.argv[2]).resolve()
display = os.environ.get("DISPLAY", ":99")
seconds = int(os.environ.get("RECORD_SECONDS", "11"))

def run(cmd, check=True):
    return subprocess.run(cmd, check=check, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

# Kill any stale Calc process from a previous item.
subprocess.run(["pkill", "-f", "soffice"], check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1)

# Open the actual .xlsx file in Calc.
proc = subprocess.Popen(
    ["libreoffice", "--nologo", "--nodefault", "--norestore", "--nofirststartwizard", str(xlsx)],
    env={**os.environ, "DISPLAY": display},
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)

time.sleep(4)

# Maximize the active Calc window.
subprocess.run(["xdotool", "search", "--onlyvisible", "--class", "libreoffice", "windowactivate", "--sync"], check=False)
subprocess.run(["xdotool", "key", "alt+F10"], check=False)
time.sleep(1)

# Move to A5, then demonstrate data entry by editing the first visible data cell.
# The workbook already contains realistic data, so the user sees genuine cells.
# We edit the first data row to reinforce the "enter data" visual.
subprocess.run(["xdotool", "key", "ctrl+home"], check=False)
time.sleep(.4)
subprocess.run(["xdotool", "key", "down"], check=False)
subprocess.run(["xdotool", "key", "down"], check=False)
subprocess.run(["xdotool", "key", "down"], check=False)
subprocess.run(["xdotool", "key", "down"], check=False)
subprocess.run(["xdotool", "key", "right"], check=False)
subprocess.run(["xdotool", "key", "f2"], check=False)
time.sleep(.2)
subprocess.run(["xdotool", "key", "ctrl+a"], check=False)
subprocess.run(["xdotool", "type", "--delay", "70", "92"], check=False)
subprocess.run(["xdotool", "key", "Return"], check=False)
time.sleep(.7)

# Select the genuine Formula cell. Its column is dataset column count + 2.
# The orchestrator supplies the exact 1-based column number.
formula_col = int(os.environ.get("FORMULA_COL", "5"))
subprocess.run(["xdotool", "key", "ctrl+home"], check=False)
for _ in range(4):
    subprocess.run(["xdotool", "key", "down"], check=False)
for _ in range(formula_col - 1):
    subprocess.run(["xdotool", "key", "right"], check=False)
subprocess.run(["xdotool", "key", "f2"], check=False)
time.sleep(.2)

# The formula is supplied through stdin by the orchestrator via environment.
formula = os.environ.get("EXCEL_FORMULA", "")
if formula:
    subprocess.run(["xdotool", "key", "ctrl+a"], check=False)
    subprocess.run(["xdotool", "type", "--delay", "35", formula], check=False)
    subprocess.run(["xdotool", "key", "Return"], check=False)

# Hold on the result for the remainder of the capture.
time.sleep(max(1, seconds - 6))

# Close Calc.
subprocess.run(["xdotool", "key", "alt+F4"], check=False)
time.sleep(1)
subprocess.run(["pkill", "-f", "soffice"], check=False)
print(video)
