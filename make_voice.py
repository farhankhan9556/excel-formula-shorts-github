#!/usr/bin/env python3
from __future__ import annotations
import asyncio
import sys
from pathlib import Path
import edge_tts

VOICES = [
    "en-US-AndrewMultilingualNeural",
    "en-US-GuyNeural",
    "en-US-ChristopherNeural",
    "en-US-EmmaMultilingualNeural",
]

async def synthesize(text: str, output: Path):
    output.parent.mkdir(parents=True, exist_ok=True)
    last = None
    for voice in VOICES:
        for attempt in range(1, 4):
            try:
                print(f"TTS: {voice}, attempt {attempt}")
                output.unlink(missing_ok=True)
                communicate = edge_tts.Communicate(
                    text=text,
                    voice=voice,
                    rate="+5%",
                    volume="+0%",
                    pitch="+0Hz",
                )
                await communicate.save(str(output))
                if output.exists() and output.stat().st_size >= 1000:
                    print(f"TTS OK: {output}")
                    return
                raise RuntimeError("TTS produced an invalid/empty file")
            except Exception as exc:
                last = exc
                print(f"TTS failed: {type(exc).__name__}: {exc}")
                await asyncio.sleep(2 * attempt)
    raise RuntimeError(f"All Edge TTS voices failed. Last error: {last}")

def main():
    if len(sys.argv) != 3:
        raise SystemExit('Usage: python make_voice.py "Text" output.mp3')
    text = sys.argv[1].strip()
    if not text:
        raise SystemExit("TTS text is empty")
    asyncio.run(synthesize(text, Path(sys.argv[2])))

if __name__ == "__main__":
    main()
