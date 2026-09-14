#!/usr/bin/env python3
from __future__ import annotations
import argparse,json,os,subprocess,sys,time
from datetime import date,datetime
from pathlib import Path
from openpyxl.formula.translate import Translator
from formula_library import FORMULAS,get_daily_formulas
ROOT=Path(__file__).resolve().parent; OUT=ROOT/'output'

def sh(cmd,env=None):
    cmd=[str(x) for x in cmd]; print('$',' '.join(cmd)); r=subprocess.run(cmd,env=env,check=False)
    if r.returncode: raise RuntimeError(f'Command failed ({r.returncode}): {" ".join(cmd)}')

def translate_formula(formula):
    if not formula or not formula.startswith('='): return formula
    try: return Translator(formula,origin='A2').translate_formula('A5')
    except Exception as e: raise RuntimeError(f'Could not translate formula {formula!r}: {e}')

def start_xvfb():
    subprocess.run(['pkill','-f','Xvfb :99'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,check=False)
    p=subprocess.Popen(['Xvfb',':99','-screen','0','1365x900x24','-ac'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    time.sleep(2)
    if p.poll() is not None: raise RuntimeError('Xvfb :99 failed to start')
    return p

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--date'); ap.add_argument('--count',type=int,default=3); a=ap.parse_args()
    if not 1<=a.count<=len(FORMULAS): raise SystemExit(f'count must be 1-{len(FORMULAS)}')
    run_date=datetime.strptime(a.date,'%Y-%m-%d').date() if a.date else date.today()
    OUT.mkdir(exist_ok=True)
    for p in OUT.iterdir():
        if p.is_file(): p.unlink()
    selected=get_daily_formulas(run_date,a.count); indexes=[FORMULAS.index(x) for x in selected]
    manifest={'date':run_date.isoformat(),'videos':[]}; xvfb=start_xvfb()
    try:
        for n,(idx,item) in enumerate(zip(indexes,selected),1):
            stem=f'excel_{run_date.isoformat()}_{n:02d}_{item["id"]}'; xlsx=OUT/f'{stem}.xlsx'; raw=Path('/tmp')/f'{stem}_calc.mp4'; voice=OUT/f'{stem}.mp3'; final=OUT/f'{stem}.mp4'
            for p in (xlsx,raw,voice,final): p.unlink(missing_ok=True)
            formula=translate_formula(item['formula']); print(f'\n=== VIDEO {n}: {item["title"]} ===\nFormula: {formula}')
            sh([sys.executable,ROOT/'make_workbook.py',idx,xlsx]); sh([sys.executable,ROOT/'make_voice.py',item['voice'],voice])
            env=os.environ.copy(); env.update({'DISPLAY':':99','EXCEL_FORMULA':formula,'FORMULA_COL':str(len(item['headers'])+2),'DEMO_VALUE':str(item['rows'][0][0]) if item['rows'] else '','RECORD_SECONDS':'12'})
            sh([sys.executable,ROOT/'record_calc.py',xlsx,raw],env=env)
            sh([sys.executable,ROOT/'render_video.py',raw,voice,final,item['problem'],item['title'],formula,item['result'],ROOT/'assets/learnverse_logo.png'])
            manifest['videos'].append({'file':final.name,'workbook':xlsx.name,'formula_id':item['id'],'problem':item['problem'],'title':item['title'],'formula':formula,'result':item['result'],'voice':item['voice']})
            voice.unlink(missing_ok=True); raw.unlink(missing_ok=True)
        (OUT/'manifest.json').write_text(json.dumps(manifest,indent=2),encoding='utf-8'); print(json.dumps(manifest,indent=2))
    finally:
        xvfb.terminate()
        try: xvfb.wait(timeout=5)
        except subprocess.TimeoutExpired: xvfb.kill()
        subprocess.run(['pkill','-f','soffice.bin'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,check=False)
        subprocess.run(['pkill','-f','soffice'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,check=False)
if __name__=='__main__': main()
