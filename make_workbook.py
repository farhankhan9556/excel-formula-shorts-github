#!/usr/bin/env python3
from pathlib import Path
import sys
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
from formula_library import FORMULAS
START=5

def translate_formula(formula):
    if not formula or not formula.startswith('='): return formula
    return Translator(formula,origin='A2').translate_formula('A5')

def create(item,out):
    out=Path(out); out.parent.mkdir(parents=True,exist_ok=True)
    wb=Workbook(); ws=wb.active; ws.title='Excel Tip'
    ws['A1']='Learn Verse — Excel Quick Tip'; ws['A1'].font=Font(size=18,bold=True,color='217346')
    ws['A2']=item['problem']; ws['A2'].font=Font(size=11,italic=True,color='555555'); ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=max(3,len(item['headers'])))
    for c,h in enumerate(item['headers'],1):
        cell=ws.cell(4,c,h); cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='217346'); cell.alignment=Alignment(horizontal='center')
    for r,row in enumerate(item['rows'],START):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
    fcol=len(item['headers'])+2; rcol=fcol+1
    for c,title in ((fcol,'Formula'),(rcol,'Outcome')):
        cell=ws.cell(4,c,title); cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='4472C4'); cell.alignment=Alignment(horizontal='center')
    formula=translate_formula(item['formula']); ws.cell(START,fcol,formula); ws.cell(START,rcol,item['result'])
    thin=Side(style='thin',color='D9E1F2')
    for row in ws.iter_rows(min_row=4,max_row=START+len(item['rows'])-1,min_col=1,max_col=rcol):
        for cell in row: cell.border=Border(left=thin,right=thin,top=thin,bottom=thin); cell.alignment=Alignment(vertical='center')
    for c in range(1,rcol+1):
        vals=[str(ws.cell(r,c).value or '') for r in range(1,START+len(item['rows']))]; ws.column_dimensions[get_column_letter(c)].width=max(14,min(32,max(map(len,vals))+3))
    ws.freeze_panes='A5'; ws.auto_filter.ref=f'A4:{get_column_letter(len(item["headers"]))}{START+len(item["rows"])-1}'
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'; wb.save(out)
    chk=load_workbook(out,read_only=True,data_only=False); actual=chk['Excel Tip'].cell(START,fcol).value; chk.close()
    if actual!=formula: raise RuntimeError(f'Formula verification failed: {actual!r} != {formula!r}')
    print('Workbook:',out); print('Formula:',formula); print('Expected:',item['result'])
if __name__=='__main__':
    if len(sys.argv)!=3: raise SystemExit('Usage: python make_workbook.py <index> <output.xlsx>')
    create(FORMULAS[int(sys.argv[1])],Path(sys.argv[2]))
