from __future__ import annotations
import argparse,asyncio,datetime,json,subprocess,sys,time
from pathlib import Path
import win32com.client as win32
from formula_library import get_daily_videos
from make_excel_demo import build_workbook
from excel_actions import focus_excel,click_cell,type_text,type_formula
from record_excel import record_excel
from make_voice import make as make_voice
ROOT=Path(__file__).resolve().parent; OUT=ROOT/"output"

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--date',default=''); ap.add_argument('--count',type=int,default=3); a=ap.parse_args()
    run_date=datetime.datetime.strptime(a.date,'%Y-%m-%d').date() if a.date else datetime.date.today()
    OUT.mkdir(exist_ok=True)
    for p in OUT.glob('*'): p.unlink()
    videos=get_daily_videos(run_date,a.count); manifest=[]
    for idx,v in enumerate(videos,1):
        stem=f"learnverse_{run_date}_{idx:02d}_{v['id']}"; xlsx=OUT/f"{stem}.xlsx"; raw=OUT/f"{stem}_raw.mp4"; voice=OUT/f"{stem}.mp3"; final=OUT/f"{stem}.mp4"
        excel=book=None
        try:
            excel,book=build_workbook(v,xlsx); excel.Visible=True; excel.WindowState=-4137; book.Activate(); book.Worksheets('Learn Verse').Activate(); focus_excel(excel)
            asyncio.run(make_voice(v['voice'],voice)); sheet=book.Worksheets('Learn Verse')
            rec=record_excel(excel,raw,30); time.sleep(1)
            if v['type']=='checkbox':
                for r in (4,5,7): click_cell(excel,sheet,r,2); time.sleep(.4)
            else:
                type_text(excel,sheet,4,5,v['input']); time.sleep(.8)
                type_formula(excel,sheet,6,6,v['formula']); time.sleep(3)
            book.Save(); rec.wait(timeout=45)
            if rec.returncode not in (0,None): raise RuntimeError('FFmpeg recording failed')
            subprocess.run([sys.executable,str(ROOT/'caption_video.py'),str(raw),str(voice),str(final),json.dumps(v['captions'])],check=True)
            manifest.append({'video':final.name,'workbook':xlsx.name,'id':v['id'],'title':v['title']})
        finally:
            try:
                if book: book.Close(SaveChanges=False)
            except: pass
            try:
                if excel: excel.Quit()
            except: pass
            raw.unlink(missing_ok=True); voice.unlink(missing_ok=True)
    (OUT/'manifest.json').write_text(json.dumps({'date':str(run_date),'videos':manifest},indent=2))
    print(json.dumps(manifest,indent=2))
if __name__=='__main__': main()
