from __future__ import annotations

from pathlib import Path
import win32com.client as win32
from win32com.client import constants


def rgb(r, g, b):
    return r + (g * 256) + (b * 65536)


def apply_border(rng):
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
    excel.ScreenUpdating = False

    book = excel.Workbooks.Add()
    sheet = book.Worksheets(1)
    sheet.Name = "Learn Verse"

    # ---------------------------------------------------------
    # PAGE / VIEW
    # ---------------------------------------------------------
    sheet.Activate()
    excel.ActiveWindow.DisplayGridlines = False
    excel.ActiveWindow.Zoom = 95

    # ---------------------------------------------------------
    # COLUMN WIDTHS
    # ---------------------------------------------------------
    widths = {
        "A": 7,
        "B": 20,
        "C": 18,
        "D": 16,
        "E": 16,
        "F": 22,
        "G": 18,
    }

    for col, width in widths.items():
        sheet.Columns(col).ColumnWidth = width

    # ---------------------------------------------------------
    # BRANDING HEADER
    # ---------------------------------------------------------
    sheet.Range("A1:G1").Merge()
    title = sheet.Range("A1")
    title.Value = "LEARN VERSE  |  EXCEL QUICK TIP"
    title.Font.Name = "Aptos Display"
    title.Font.Size = 20
    title.Font.Bold = True
    title.Font.Color = rgb(255, 255, 255)
    title.Interior.Color = rgb(31, 78, 121)
    title.HorizontalAlignment = constants.xlCenter
    title.VerticalAlignment = constants.xlCenter
    sheet.Rows(1).RowHeight = 36

    sheet.Range("A2:G2").Merge()
    subtitle = sheet.Range("A2")
    subtitle.Value = video.get("title", video.get("id", "Excel Tip")).upper()
    subtitle.Font.Name = "Aptos"
    subtitle.Font.Size = 11
    subtitle.Font.Bold = True
    subtitle.Font.Color = rgb(31, 78, 121)
    subtitle.HorizontalAlignment = constants.xlCenter
    subtitle.VerticalAlignment = constants.xlCenter
    sheet.Rows(2).RowHeight = 24

    # ---------------------------------------------------------
    # DATA TABLE
    # ---------------------------------------------------------
    headers = list(video.get("headers", []))
    rows = list(video.get("rows", []))

    # Keep the visible table within A:G.
    max_cols = min(len(headers), 7)
    headers = headers[:max_cols]

    # Header row
    for c, header in enumerate(headers, 1):
        cell = sheet.Cells(4, c)
        cell.Value = header
        cell.Font.Name = "Aptos"
        cell.Font.Size = 10
        cell.Font.Bold = True
        cell.Font.Color = rgb(255, 255, 255)
        cell.Interior.Color = rgb(47, 117, 181)
        cell.HorizontalAlignment = constants.xlCenter
        cell.VerticalAlignment = constants.xlCenter

    sheet.Rows(4).RowHeight = 28

    # Data rows
    start_row = 5

    for r_index, row in enumerate(rows, start_row):
        for c_index, value in enumerate(row[:max_cols], 1):
            cell = sheet.Cells(r_index, c_index)
            cell.Value = value
            cell.Font.Name = "Aptos"
            cell.Font.Size = 10
            cell.VerticalAlignment = constants.xlCenter

            if c_index == 1:
                cell.Font.Bold = True

            if r_index % 2 == 1:
                cell.Interior.Color = rgb(242, 246, 250)
            else:
                cell.Interior.Color = rgb(255, 255, 255)

        sheet.Rows(r_index).RowHeight = 25

    # ---------------------------------------------------------
    # EXCEL TABLE
    # ---------------------------------------------------------
    if rows and headers:
        last_row = start_row + len(rows) - 1
        last_col = len(headers)

        table_range = sheet.Range(
            sheet.Cells(4, 1),
            sheet.Cells(last_row, last_col)
        )

        try:
            table = sheet.ListObjects.Add(
                constants.xlSrcRange,
                table_range,
                None,
                constants.xlYes
            )
            table.Name = "LearnVerseTable"

            try:
                table.TableStyle = "TableStyleMedium2"
            except Exception:
                pass

        except Exception:
            apply_border(table_range)

    # ---------------------------------------------------------
    # DEMONSTRATION AREA
    # ---------------------------------------------------------
    demo_top = start_row + len(rows) + 2

    sheet.Range(
        sheet.Cells(demo_top, 1),
        sheet.Cells(demo_top, 7)
    ).Merge()

    demo_header = sheet.Cells(demo_top, 1)
    demo_header.Value = "FORMULA DEMONSTRATION"
    demo_header.Font.Name = "Aptos Display"
    demo_header.Font.Size = 13
    demo_header.Font.Bold = True
    demo_header.Font.Color = rgb(255, 255, 255)
    demo_header.Interior.Color = rgb(31, 78, 121)
    demo_header.HorizontalAlignment = constants.xlCenter
    demo_header.VerticalAlignment = constants.xlCenter
    sheet.Rows(demo_top).RowHeight = 28

    # ---------------------------------------------------------
    # FORMULA TYPE
    # ---------------------------------------------------------
    if video.get("type") == "formula":
        input_row = demo_top + 2
        formula_row = demo_top + 4

        labels = [
            ("E", "INPUT"),
            ("F", "FORMULA"),
            ("G", "RESULT"),
        ]

        for col, text in labels:
            cell = sheet.Range(f"{col}{input_row}")
            cell.Value = text
            cell.Font.Bold = True
            cell.Font.Size = 10
            cell.Font.Color = rgb(255, 255, 255)
            cell.Interior.Color = rgb(47, 117, 181)
            cell.HorizontalAlignment = constants.xlCenter

        input_cell = sheet.Range(f"E{input_row + 1}")
        input_cell.Value = video.get("input", "")
        input_cell.Font.Bold = True
        input_cell.Font.Size = 12
        input_cell.Interior.Color = rgb(255, 242, 204)
        input_cell.HorizontalAlignment = constants.xlCenter

        formula_cell = sheet.Range(f"F{formula_row}")
        formula_cell.Value = video.get("formula", "")
        formula_cell.Font.Name = "Consolas"
        formula_cell.Font.Size = 11
        formula_cell.Font.Bold = True
        formula_cell.Interior.Color = rgb(221, 235, 247)
        formula_cell.HorizontalAlignment = constants.xlCenter

        result_cell = sheet.Range(f"G{formula_row}")

        formula = video.get("formula", "")
        if formula:
            result_cell.Formula = formula

        result_cell.Font.Size = 13
        result_cell.Font.Bold = True
        result_cell.Interior.Color = rgb(226, 239, 218)
        result_cell.HorizontalAlignment = constants.xlCenter

        # Add labels for clarity.
        sheet.Range(f"E{input_row + 2}:G{input_row + 2}").Merge()
        note = sheet.Range(f"E{input_row + 2}")
        note.Value = "Type → Apply Formula → See Result"
        note.Font.Size = 10
        note.Font.Italic = True
        note.Font.Color = rgb(89, 89, 89)
        note.HorizontalAlignment = constants.xlCenter

        apply_border(
            sheet.Range(
                f"E{input_row}",
                f"G{formula_row}"
            )
        )

    # ---------------------------------------------------------
    # CHECKBOX TYPE
    # ---------------------------------------------------------
    else:
        status_header_row = demo_top + 2

        sheet.Range(
            f"A{status_header_row}:E{status_header_row}"
        ).Merge()

        status_header = sheet.Range(f"A{status_header_row}")
        status_header.Value = "TASK STATUS"
        status_header.Font.Bold = True
        status_header.Font.Color = rgb(255, 255, 255)
        status_header.Interior.Color = rgb(47, 117, 181)
        status_header.HorizontalAlignment = constants.xlCenter

        # Real Form Control checkboxes.
        for index, _row in enumerate(rows, start_row):
            link = f"D{index}"

            sheet.Cells(index, 4).Value = False
            sheet.Cells(index, 4).NumberFormat = "General"

            sheet.Cells(index, 5).Formula = (
                f'=IF({link},"DONE","PENDING")'
            )

            status_cell = sheet.Cells(index, 5)
            status_cell.Font.Bold = True
            status_cell.HorizontalAlignment = constants.xlCenter

            try:
                checkbox = sheet.Shapes.AddFormControl(
                    8,
                    sheet.Cells(index, 2).Left + 3,
                    sheet.Cells(index, 2).Top + 3,
                    18,
                    18,
                )

                checkbox.Name = f"LV_Check_{index}"
                checkbox.ControlFormat.LinkedCell = link
                checkbox.TextFrame.Characters().Text = ""

            except Exception:
                pass

        sheet.Range(f"A{status_header_row + 1}:E{status_header_row + 2}").Merge()
        note = sheet.Range(f"A{status_header_row + 1}")
        note.Value = "Click the checkbox → Excel updates the status automatically"
        note.Font.Italic = True
        note.Font.Color = rgb(89, 89, 89)
        note.HorizontalAlignment = constants.xlCenter

    # ---------------------------------------------------------
    # FREEZE PANES
    # ---------------------------------------------------------
    try:
        sheet.Range("A5").Select()
        excel.ActiveWindow.FreezePanes = True
    except Exception:
        pass

    # ---------------------------------------------------------
    # PAGE SETUP
    # ---------------------------------------------------------
    try:
        sheet.PageSetup.Orientation = constants.xlLandscape
        sheet.PageSetup.Zoom = False
        sheet.PageSetup.FitToPagesWide = 1
        sheet.PageSetup.FitToPagesTall = False
    except Exception:
        pass

    # ---------------------------------------------------------
    # SAVE
    # ---------------------------------------------------------
    sheet.Activate()
    excel.WindowState = constants.xlMaximized

    book.SaveAs(
        str(path),
        FileFormat=51
    )

    return excel, book