from __future__ import annotations
import json,sys,tempfile,subprocess
from pathlib import Path
from caption_renderer import make_card

def main():
    if len(sys.argv)!=5: raise SystemExit("Usage: caption_video.py raw voice final steps_json")
    raw,voice,final=map(Path,sys.argv[1:4]); steps=json.loads(sys.argv[4])
    with tempfile.TemporaryDirectory() as td:
        td=Path(td); inputs=[]
        for i,step in enumerate(steps):
            p=td/f"cap{i}.png"; make_card(step[2],p); inputs.append(p)
        args=["ffmpeg","-hide_banner","-loglevel","warning","-y","-i",str(raw),"-i",str(voice)]
        for p in inputs: args += ["-loop","1","-i",str(p)]
        filters=["[0:v]setsar=1[base]"]; prev="base"
        for i,(start,end,_) in enumerate(steps):
            nxt=f"v{i}"; filters += [f"[{2+i}:v]format=rgba[c{i}]",f"[{prev}][c{i}]overlay=0:0:enable='between(t,{start},{end})'[{nxt}]"]; prev=nxt
        args += ["-filter_complex",";".join(filters),"-map",f"[{prev}]","-map","1:a","-t","30","-c:v","libx264","-preset","veryfast","-crf","20","-pix_fmt","yuv420p","-c:a","aac","-b:a","128k","-movflags","+faststart",str(final)]
        r=subprocess.run(args,capture_output=True,text=True)
        if r.returncode: raise RuntimeError(r.stderr[-4000:])
if __name__=='__main__': main()
