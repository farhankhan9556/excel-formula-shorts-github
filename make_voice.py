from __future__ import annotations
import asyncio,sys
from pathlib import Path
import edge_tts

VOICES=["en-US-AndrewMultilingualNeural","en-US-ChristopherNeural","en-US-GuyNeural"]
async def make(text,out):
    last=None
    for voice in VOICES:
        for attempt in range(3):
            try:
                Path(out).unlink(missing_ok=True)
                await edge_tts.Communicate(text=text,voice=voice,rate="+5%",pitch="+0Hz").save(str(out))
                if Path(out).exists() and Path(out).stat().st_size>1000:
                    return
            except Exception as e:
                last=e
                await asyncio.sleep(1+attempt)
    raise RuntimeError(f"TTS failed: {last}")
