from __future__ import annotations
import argparse,datetime,json,subprocess,sys,time,asyncio
from pathlib import Path
import win32com.client as win32
from formula_library import get_daily_videos
from make_excel_demo import build_workbook
from record_excel import run_checkbox,run_formula
from make_voice import make as make_voice

ROOT=Path(__file__).resolve().parent
OUT=ROOT/"output"

def create_excel(video,path):
    build_workbook(video,path)
    excel=win32.DispatchEx("Excel.Application")
    excel.Visible=True
    excel.DisplayAlerts=False
    book=excel.Workbooks.Open(str(path.resolve()))
    sheet=book.Worksheets("Learn Verse")
    excel.WindowState=-4137
    # Remove old controls and add real Excel Form Control checkboxes.
    for shape in list(sheet.Shapes):
        try: shape.Delete()
        except: pass
    if video.get("type")=="checkbox":
        for r in range(4,13):
            cell=sheet.Cells(r,2)
            cb=sheet.Shapes.AddFormControl(8,cell.Left+4,cell.Top+3,18,18)
            cb.Name=f"LV_Check_{r}"
            cb.ControlFormat.Value=0
    book.Save()
    return excel,book

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--date",default="")
    ap.add_argument("--count",type=int,default=3)
    args=ap.parse_args()
    run_date=datetime.datetime.strptime(args.date,"%Y-%m-%d").date() if args.date else datetime.date.today()
    OUT.mkdir(exist_ok=True)
    for p in OUT.iterdir():
        if p.is_file(): p.unlink()
    selected=get_daily_videos(run_date,args.count)
    manifest={"date":str(run_date),"videos":[]}
    for i,video in enumerate(selected,1):
        stem=f"learnverse_{run_date}_{i:02d}_{video['id']}"
        xlsx=OUT/f"{stem}.xlsx"; raw=OUT/f"{stem}_raw.mp4"; voice=OUT/f"{stem}.mp3"; final=OUT/f"{stem}.mp4"
        excel=book=None
        try:
            build_workbook(video,xlsx)
            excel,book=create_excel(video,xlsx)
            asyncio.run(make_voice(video["voice"], voice))
            if video.get("type")=="checkbox":
                run_checkbox(excel,book,video,raw)
            else:
                run_formula(excel,book,video,raw)
            book.Close(SaveChanges=False); book=None
            excel.Quit(); excel=None
            # Final edit with captions is done by caption_video.py
            subprocess.run([sys.executable,str(ROOT/"caption_video.py"),str(raw),str(voice),str(final),json.dumps(video["caption_steps"])],check=True)
            raw.unlink(missing_ok=True); voice.unlink(missing_ok=True)
            manifest["videos"].append({"file":final.name,"workbook":xlsx.name,"id":video["id"],"title":video["title"]})
        finally:
            if book:
                try: book.Close(SaveChanges=False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
    (OUT/"manifest.json").write_text(json.dumps(manifest,indent=2))
    print(json.dumps(manifest,indent=2))
if __name__=="__main__": main()
