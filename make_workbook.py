from pathlib import Path
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from formula_library import FORMULAS

def create_workbook(item, output):
    wb = Workbook()
    ws = wb.active
    ws.title = "Excel Tip"

    ws["A1"] = "Learn Verse — Excel Quick Tip"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A2"] = item["problem"]
    ws["A2"].font = Font(size=11, italic=True)

    headers = item["headers"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="217346")
        cell.alignment = Alignment(horizontal="center")

    for r, row in enumerate(item["rows"], 5):
        for c, value in enumerate(row, 1):
            ws.cell(row=r, column=c, value=value)

    # Formula/result area to the right, using a genuine Excel formula.
    formula_col = len(headers) + 2
    result_col = formula_col + 1
    ws.cell(row=4, column=formula_col, value="Formula")
    ws.cell(row=4, column=result_col, value="Outcome")
    for c in (formula_col, result_col):
        ws.cell(row=4, column=c).font = Font(bold=True, color="FFFFFF")
        ws.cell(row=4, column=c).fill = PatternFill("solid", fgColor="4472C4")
        ws.cell(row=4, column=c).alignment = Alignment(horizontal="center")

    # Convert formula references from old row numbering to workbook row numbering.
    # The video types the original formula into Calc; the workbook itself uses
    # corresponding workbook row references.
    formula = item["formula"]
    formula_for_workbook = formula.replace("B2", "B5").replace("A2", "A5").replace("C2", "C5").replace("D2", "D5").replace("E2", "E5")
    formula_for_workbook = formula_for_workbook.replace("B4", f"B{4+len(item['rows'])}").replace("A4", f"A{4+len(item['rows'])}").replace("C4", f"C{4+len(item['rows'])}").replace("D4", f"D{4+len(item['rows'])}").replace("E4", f"E{4+len(item['rows'])}")
    formula_for_workbook = formula_for_workbook.replace("B5", f"B{5+len(item['rows'])-1}").replace("A5", f"A{5+len(item['rows'])-1}").replace("C5", f"C{5+len(item['rows'])-1}").replace("D5", f"D{5+len(item['rows'])-1}").replace("E5", f"E{5+len(item['rows'])-1}")

    ws.cell(row=5, column=formula_col, value=formula)
    ws.cell(row=5, column=result_col, value=item["result"])

    thin = Side(style="thin", color="D9E1F2")
    for row in ws.iter_rows(min_row=4, max_row=max(5, 4+len(item["rows"])), min_col=1, max_col=result_col):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center")

    for col in range(1, result_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = max(14, min(28, max(len(str(ws.cell(r, col).value or "")) for r in range(1, 6)) + 3))

    ws.freeze_panes = "A5"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(output)

if __name__ == "__main__":
    # usage: python make_workbook.py <index> <output>
    idx = int(sys.argv[1])
    out = Path(sys.argv[2])
    create_workbook(FORMULAS[idx], out)
    print(out)
