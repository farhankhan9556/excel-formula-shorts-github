from __future__ import annotations
import asyncio, sys
from pathlib import Path
import edge_tts

VOICES=["en-US-AndrewMultilingualNeural","en-US-GuyNeural","en-US-ChristopherNeural","en-US-EmmaMultilingualNeural"]

async def make(text,out):
    last=None
    for voice in VOICES:
        for attempt in range(1,4):
            try:
                Path(out).unlink(missing_ok=True)
                await edge_tts.Communicate(text=text,voice=voice,rate="+3%",pitch="+0Hz").save(str(out))
                if Path(out).exists() and Path(out).stat().st_size>1000:
                    print("TTS OK",voice,out); return
            except Exception as e:
                last=e
                await asyncio.sleep(attempt)
    raise RuntimeError(f"TTS failed: {last}")

if __name__=="__main__":
    if len(sys.argv)!=3: raise SystemExit('Usage: python make_voice.py "text" output.mp3')
    asyncio.run(make(sys.argv[1],sys.argv[2]))
