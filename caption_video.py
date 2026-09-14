from __future__ import annotations
import json,sys,tempfile,subprocess
from pathlib import Path
from caption_renderer import make_card

def main():
    if len(sys.argv)!=5: raise SystemExit("Usage: caption_video.py raw.mp4 voice.mp3 final.mp4 steps_json")
    raw,voice,final=map(Path,sys.argv[1:4]); steps=json.loads(sys.argv[4])
    with tempfile.TemporaryDirectory() as td:
        td=Path(td)
        inputs=[]
        for n,(_,_,text) in enumerate(steps):
            p=td/f"c{n}.png"; make_card(text,p); inputs.append(p)
        # One looping image per caption; reference-style captions are timed and kept low.
        args=["ffmpeg","-hide_banner","-loglevel","warning","-y","-i",str(raw),"-i",str(voice)]
        for p in inputs: args += ["-loop","1","-i",str(p)]
        filters=["[0:v]scale=1080:1920:force_original_aspect_ratio=decrease,pad=1080:1920:(ow-iw)/2:(oh-ih)/2,setsar=1[base]"]
        prev="base"
        for i,(start,end,text) in enumerate(steps):
            nxt=f"v{i}"
            filters.append(f"[{2+i}:v]format=rgba[c{i}]")
            filters.append(f"[{prev}][c{i}]overlay=0:0:enable='between(t,{start},{end})'[{nxt}]")
            prev=nxt
        filt=";".join(filters)
        args += ["-filter_complex",filt,"-map",f"[{prev}]","-map","1:a","-t","30","-c:v","libx264","-preset","veryfast","-crf","20","-pix_fmt","yuv420p","-c:a","aac","-b:a","128k","-movflags","+faststart",str(final)]
        r=subprocess.run(args,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
        print(r.stdout)
        if r.returncode: raise RuntimeError("caption render failed")
if __name__=="__main__": main()
