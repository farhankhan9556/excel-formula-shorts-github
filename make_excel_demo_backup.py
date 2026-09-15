from __future__ import annotations
from pathlib import Path
import win32com.client as win32
from win32com.client import constants

def build_workbook(video,path):
    path=Path(path).resolve(); path.parent.mkdir(parents=True,exist_ok=True)
    excel=win32.DispatchEx("Excel.Application"); excel.Visible=False; excel.DisplayAlerts=False
    book=excel.Workbooks.Add(); sheet=book.Worksheets(1); sheet.Name="Learn Verse"
    sheet.Range("A1:G1").Merge(); sheet.Range("A1").Value="Learn Verse — Excel Quick Tip"
    sheet.Range("A1").Font.Size=18; sheet.Range("A1").Font.Bold=True
    headers=video["headers"]
    for c,h in enumerate(headers,1):
        cell=sheet.Cells(3,c); cell.Value=h; cell.Font.Bold=True; cell.Font.Color=0xFFFFFF; cell.Interior.Color=0x00A5FF
    for r,row in enumerate(video["rows"],4):
        for c,v in enumerate(row,1): sheet.Cells(r,c).Value=v
    for c in range(1,len(headers)+1): sheet.Columns(c).ColumnWidth=18
    sheet.Rows("3:20").RowHeight=25
    if video["type"]=="formula":
        sheet.Cells(3,5).Value="Input"; sheet.Cells(3,6).Value="Formula"; sheet.Cells(3,7).Value="Result"
        for c in (5,6,7): sheet.Cells(3,c).Font.Bold=True; sheet.Cells(3,c).Font.Color=0xFFFFFF; sheet.Cells(3,c).Interior.Color=0x00A5FF
        sheet.Cells(4,5).Value=video["input"]
        sheet.Cells(6,6).Value=video["formula"]
        sheet.Cells(6,7).Formula=video["formula"]
        sheet.Cells(4,5).Interior.Color=0x00FFFF
        sheet.Cells(6,7).Interior.Color=0xCCFFCC
        sheet.Columns(6).ColumnWidth=34; sheet.Columns(7).ColumnWidth=18
    else:
        # Real Excel Form Control checkboxes linked to cells.
        for r in range(4,4+len(video["rows"])):
            link=f"D{r}"
            sheet.Cells(r,4).Value=False
            sheet.Cells(r,5).Formula=f'=IF({link},"DONE","PENDING")'
            sheet.Cells(r,4).NumberFormat="General"
            shp=sheet.Shapes.AddFormControl(8,sheet.Cells(r,2).Left+3,sheet.Cells(r,2).Top+3,18,18)
            shp.Name=f"LV_Check_{r}"; shp.ControlFormat.LinkedCell=link
        sheet.Cells(3,4).Value="Linked"; sheet.Cells(3,5).Value="Status"
        for c in (4,5): sheet.Cells(3,c).Font.Bold=True; sheet.Cells(3,c).Font.Color=0xFFFFFF; sheet.Cells(3,c).Interior.Color=0x00A5FF
    sheet.Activate(); excel.WindowState=-4137
    book.SaveAs(str(path),FileFormat=51)
    return excel,book
