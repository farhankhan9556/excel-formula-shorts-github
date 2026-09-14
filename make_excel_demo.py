from __future__ import annotations
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

def build_workbook(video, path):
    path=Path(path).resolve()
    path.parent.mkdir(parents=True,exist_ok=True)
    wb=Workbook()
    ws=wb.active
    ws.title="Learn Verse"
    ws.sheet_view.showGridLines=True
    ws["A1"]="Learn Verse — Excel Quick Tip"
    ws["A1"].font=Font(size=16,bold=True,color="217346")
    for c,h in enumerate(video["headers"],1):
        cell=ws.cell(3,c,h)
        cell.font=Font(bold=True,color="FFFFFF")
        cell.fill=PatternFill("solid",fgColor="FFC000")
        cell.alignment=Alignment(horizontal="center")
    for r,row in enumerate(video["rows"],4):
        for c,v in enumerate(row,1):
            ws.cell(r,c,"" if isinstance(v,bool) else v)
    if video.get("type")=="formula":
        ws["F3"]="Formula"; ws["G3"]="Result"
        for addr in ("F3","G3"):
            ws[addr].font=Font(bold=True,color="FFFFFF")
            ws[addr].fill=PatternFill("solid",fgColor="4472C4")
        ws["F4"]=video["formula"]; ws["G4"]=video["result"]
    widths={"A":20,"B":16,"C":18,"D":16,"E":3,"F":40,"G":16}
    for col,w in widths.items(): ws.column_dimensions[col].width=w
    for r in range(4,15): ws.row_dimensions[r].height=24
    wb.save(path)
