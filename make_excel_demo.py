from __future__ import annotations

from pathlib import Path
import win32com.client as win32
from win32com.client import constants


def rgb(r, g, b):
    return r + g * 256 + b * 65536


def border(rng):
    for i in range(7, 13):
        try:
            rng.Borders(i).LineStyle = 1
            rng.Borders(i).Weight = 2
        except Exception:
            pass


def build_workbook(video, path):
    path = Path(path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = True

    book = excel.Workbooks.Add()
    sheet = book.Worksheets(1)
    sheet.Name = "Learn Verse"
    sheet.Activate()

    try:
        excel.ActiveWindow.DisplayGridlines = False
        excel.ActiveWindow.Zoom = 95
    except Exception:
        pass

    widths = {"A": 18, "B": 18, "C": 16, "D": 18, "E": 18, "F": 28, "G": 20}
    for c, w in widths.items():
        sheet.Columns(c).ColumnWidth = w

    sheet.Range("A1:G1").Merge()
    t = sheet.Range("A1")
    t.Value = "LEARN VERSE  |  EXCEL QUICK TIP"
    t.Font.Name = "Aptos Display"
    t.Font.Size = 20
    t.Font.Bold = True
    t.Font.Color = rgb(255, 255, 255)
    t.Interior.Color = rgb(31, 78, 121)
    t.HorizontalAlignment = constants.xlCenter
    t.VerticalAlignment = constants.xlCenter
    sheet.Rows(1).RowHeight = 38

    sheet.Range("A2:G2").Merge()
    s = sheet.Range("A2")
    s.Value = video["title"].upper()
    s.Font.Name = "Aptos"
    s.Font.Size = 12
    s.Font.Bold = True
    s.Font.Color = rgb(31, 78, 121)
    s.HorizontalAlignment = constants.xlCenter
    s.VerticalAlignment = constants.xlCenter
    sheet.Rows(2).RowHeight = 25

    headers = video["headers"]
    for c, h in enumerate(headers, 1):
        cell = sheet.Cells(4, c)
        cell.Value = h
        cell.Font.Name = "Aptos"
        cell.Font.Size = 10
        cell.Font.Bold = True
        cell.Font.Color = rgb(255, 255, 255)
        cell.Interior.Color = rgb(47, 117, 181)
        cell.HorizontalAlignment = constants.xlCenter
        cell.VerticalAlignment = constants.xlCenter
    sheet.Rows(4).RowHeight = 28

    rows = video["rows"]
    for r, row in enumerate(rows, 5):
        for c, value in enumerate(row, 1):
            cell = sheet.Cells(r, c)
            cell.Value = value
            cell.Font.Name = "Aptos"
            cell.Font.Size = 10
            cell.VerticalAlignment = constants.xlCenter
            if c == 1:
                cell.Font.Bold = True
            if r % 2:
                cell.Interior.Color = rgb(242, 246, 250)
        sheet.Rows(r).RowHeight = 25

    last_row = 4 + len(rows)
    last_col = len(headers)
    try:
        table = sheet.ListObjects.Add(
            constants.xlSrcRange,
            sheet.Range(sheet.Cells(4, 1), sheet.Cells(last_row, last_col)),
            None,
            constants.xlYes,
        )
        table.Name = "LearnVerseTable"
        table.TableStyle = "TableStyleMedium2"
    except Exception:
        border(sheet.Range(sheet.Cells(4, 1), sheet.Cells(last_row, last_col)))

    demo_top = last_row + 2
    sheet.Range(f"A{demo_top}:G{demo_top}").Merge()
    dh = sheet.Range(f"A{demo_top}")
    dh.Value = "LIVE FORMULA DEMONSTRATION"
    dh.Font.Name = "Aptos Display"
    dh.Font.Size = 13
    dh.Font.Bold = True
    dh.Font.Color = rgb(255, 255, 255)
    dh.Interior.Color = rgb(31, 78, 121)
    dh.HorizontalAlignment = constants.xlCenter
    dh.VerticalAlignment = constants.xlCenter
    sheet.Rows(demo_top).RowHeight = 30

    if video["type"] == "formula":
        label_row = demo_top + 2
        input_row = label_row + 1
        formula_row = label_row + 3

        for col, text in [("E", "INPUT"), ("F", "FORMULA"), ("G", "RESULT")]:
            cell = sheet.Range(f"{col}{label_row}")
            cell.Value = text
            cell.Font.Bold = True
            cell.Font.Color = rgb(255, 255, 255)
            cell.Interior.Color = rgb(47, 117, 181)
            cell.HorizontalAlignment = constants.xlCenter

        inp = sheet.Range(video["input_cell"])
        inp.Value = video.get("input", "")
        inp.Font.Bold = True
        inp.Font.Size = 13
        inp.Interior.Color = rgb(255, 242, 204)
        inp.HorizontalAlignment = constants.xlCenter
        inp.NumberFormat = "General"

        fcell = sheet.Range(video["formula_cell"])
        fcell.Value = video["formula"]
        fcell.Font.Name = "Consolas"
        fcell.Font.Size = 10
        fcell.Font.Bold = True
        fcell.Interior.Color = rgb(221, 235, 247)
        fcell.HorizontalAlignment = constants.xlCenter

        result = sheet.Range(video["result_cell"])
        result.Formula = video["formula"]
        result.Font.Size = 13
        result.Font.Bold = True
        result.Interior.Color = rgb(226, 239, 218)
        result.HorizontalAlignment = constants.xlCenter

        sheet.Range(f"E{input_row+2}:G{input_row+2}").Merge()
        note = sheet.Range(f"E{input_row+2}")
        note.Value = "TYPE  →  APPLY FORMULA  →  SEE RESULT"
        note.Font.Italic = True
        note.Font.Color = rgb(89, 89, 89)
        note.HorizontalAlignment = constants.xlCenter
        border(sheet.Range(f"E{label_row}:G{formula_row}"))

    else:
        status_row = demo_top + 2
        sheet.Range(f"A{status_row}:D{status_row}").Merge()
        sh = sheet.Range(f"A{status_row}")
        sh.Value = "CLICK THE REAL EXCEL CHECKBOXES"
        sh.Font.Bold = True
        sh.Font.Color = rgb(255, 255, 255)
        sh.Interior.Color = rgb(47, 117, 181)
        sh.HorizontalAlignment = constants.xlCenter

        for r in range(5, 5 + len(rows)):
            # Column C stores TRUE/FALSE and D displays the status.
            sheet.Cells(r, 3).Value = False
            sheet.Cells(r, 3).NumberFormat = "General"
            sheet.Cells(r, 4).Formula = f'=IF(C{r},"DONE","PENDING")'
            sheet.Cells(r, 4).Font.Bold = True
            sheet.Cells(r, 4).HorizontalAlignment = constants.xlCenter
            try:
                cb = sheet.Shapes.AddFormControl(
                    8,
                    sheet.Cells(r, 3).Left + 2,
                    sheet.Cells(r, 3).Top + 2,
                    18,
                    18,
                )
                cb.Name = f"LV_Check_{r}"
                cb.ControlFormat.LinkedCell = f"C{r}"
                cb.TextFrame.Characters().Text = ""
            except Exception:
                pass

        sheet.Range(f"A{status_row+1}:D{status_row+2}").Merge()
        note = sheet.Range(f"A{status_row+1}")
        note.Value = "Click → TRUE → DONE"
        note.Font.Italic = True
        note.Font.Color = rgb(89, 89, 89)
        note.HorizontalAlignment = constants.xlCenter
        border(sheet.Range(f"A{5}:D{4+len(rows)}"))

    try:
        sheet.Range("A5").Select()
        excel.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    try:
        sheet.PageSetup.Orientation = constants.xlLandscape
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
    except Exception:
        pass

    excel.WindowState = constants.xlMaximized
    excel.ScreenUpdating = True
    excel.Visible = True
    book.SaveAs(str(path), FileFormat=51)
    return excel, book
