#!/usr/bin/env python3
from __future__ import annotations
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
from formula_library import FORMULAS

def workbook_formula(item):
    # Formula library uses row 2 as the logical first data row.
    # In the generated workbook data begins on row 5.
    return Translator(item["formula"], origin="A2").translate_formula("A5")

def create_workbook(item, output: Path):
    wb=Workbook()
    ws=wb.active
    ws.title="Excel Tip"
    ws["A1"]="Learn Verse — Excel Quick Tip"
    ws["A1"].font=Font(size=18,bold=True,color="FFFFFF")
    ws["A1"].fill=PatternFill("solid",fgColor="217346")
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=max(6,len(item["headers"])+2))
    ws["A2"]=item["problem"]
    ws["A2"].font=Font(size=11,italic=True,color="404040")
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=max(6,len(item["headers"])+2))
    headers=item["headers"]
    formula_col=len(headers)+2
    result_col=formula_col+1
    for c,h in enumerate(headers,1):
        cell=ws.cell(4,c,h); cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="217346"); cell.alignment=Alignment(horizontal="center")
    ws.cell(4,formula_col,"Formula").font=Font(bold=True,color="FFFFFF")
    ws.cell(4,result_col,"Outcome").font=Font(bold=True,color="FFFFFF")
    ws.cell(4,formula_col).fill=PatternFill("solid",fgColor="4472C4")
    ws.cell(4,result_col).fill=PatternFill("solid",fgColor="4472C4")
    for r,row in enumerate(item["rows"],5):
        for c,value in enumerate(row,1):
            ws.cell(r,c,value)
    actual_formula=workbook_formula(item)
    ws.cell(5,formula_col,actual_formula)
    ws.cell(5,result_col,item["result"])
    thin=Side(style="thin",color="D9E1F2")
    for row in ws.iter_rows(min_row=4,max_row=4+len(item["rows"]),min_col=1,max_col=result_col):
        for cell in row:
            cell.border=Border(left=thin,right=thin,top=thin,bottom=thin)
            cell.alignment=Alignment(vertical="center")
    for col in range(1,result_col+1):
        letter=get_column_letter(col)
        values=[str(ws.cell(r,col).value or "") for r in range(1,5+len(item["rows"]))]
        width=max(14,min(32,max(len(v) for v in values)+3))
        ws.column_dimensions[letter].width=width
    ws.freeze_panes="A5"
    ws.auto_filter.ref=f"A4:{get_column_letter(result_col)}{4+len(item['rows'])}"
    ws.sheet_view.showGridLines=True
    wb.calculation.fullCalcOnLoad=True
    wb.calculation.forceFullCalc=True
    wb.calculation.calcMode="auto"
    output.parent.mkdir(parents=True,exist_ok=True)
    wb.save(output)
    print(f"Workbook created: {output}")
    print(f"Formula used in workbook: {actual_formula}")
    return actual_formula

if __name__=="__main__":
    if len(sys.argv)!=3:
        raise SystemExit("Usage: python make_workbook.py <index> <output.xlsx>")
    idx=int(sys.argv[1])
    if idx<0 or idx>=len(FORMULAS):
        raise SystemExit(f"Formula index must be 0-{len(FORMULAS)-1}")
    create_workbook(FORMULAS[idx],Path(sys.argv[2]))
