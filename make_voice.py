#!/usr/bin/env python3
from __future__ import annotations
import asyncio,sys
from pathlib import Path
import edge_tts
VOICES=['en-US-AndrewMultilingualNeural','en-US-GuyNeural','en-US-ChristopherNeural','en-US-EmmaMultilingualNeural']
async def synth(text,out):
    out.parent.mkdir(parents=True,exist_ok=True); last=None
    for voice in VOICES:
        for attempt in range(1,4):
            try:
                print(f'TTS {voice} attempt {attempt}'); out.unlink(missing_ok=True)
                await edge_tts.Communicate(text=text,voice=voice,rate='+5%',volume='+0%',pitch='+0Hz').save(str(out))
                if out.exists() and out.stat().st_size>=1000: print(f'TTS OK {out}'); return
                raise RuntimeError('TTS returned an invalid file')
            except Exception as e: last=e; print(f'TTS failed: {type(e).__name__}: {e}'); await asyncio.sleep(2*attempt)
    raise RuntimeError(f'All Edge TTS attempts failed: {last}')
def main():
    if len(sys.argv)!=3: raise SystemExit('Usage: python make_voice.py "Text" output.mp3')
    text=sys.argv[1].strip(); out=Path(sys.argv[2])
    if not text: raise SystemExit('TTS text is empty')
    asyncio.run(synth(text,out))
if __name__=='__main__': main()
