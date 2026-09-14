#!/usr/bin/env python3
"""
Reliable Edge-TTS wrapper for GitHub Actions.

The previous version failed inside make_voice.py. This version:
- uses the official edge-tts Python package interface;
- checks that an MP3 was actually created;
- retries transient network failures;
- tries several known English voices;
- writes useful diagnostics to the Actions log.
"""

import asyncio
import sys
from pathlib import Path
import edge_tts

VOICES = [
    "en-US-AndrewMultilingualNeural",   # male
    "en-US-GuyNeural",                  # male fallback
    "en-US-ChristopherNeural",          # male fallback
    "en-US-EmmaMultilingualNeural",     # final fallback
]

async def synthesize(text: str, output: Path):
    output.parent.mkdir(parents=True, exist_ok=True)
    last_error = None

    for voice in VOICES:
        for attempt in range(1, 4):
            try:
                print(f"TTS: voice={voice}, attempt={attempt}")

                if output.exists():
                    output.unlink()

                communicate = edge_tts.Communicate(
                    text=text,
                    voice=voice,
                    rate="+5%",
                    volume="+0%",
                    pitch="+0Hz",
                )

                await communicate.save(str(output))

                if output.exists() and output.stat().st_size >= 1000:
                    print(f"TTS SUCCESS: {voice} -> {output} ({output.stat().st_size} bytes)")
                    return

                raise RuntimeError("Edge TTS returned without a usable audio file.")

            except Exception as exc:
                last_error = exc
                print(f"TTS attempt failed: {type(exc).__name__}: {exc}")
                await asyncio.sleep(2 * attempt)

    raise RuntimeError(
        "All Edge-TTS attempts failed. "
        "The GitHub runner may not be able to reach Microsoft's Edge TTS service. "
        f"Last error: {last_error}"
    )

def main():
    if len(sys.argv) != 3:
        raise SystemExit(
            'Usage: python make_voice.py "Text to speak" output.mp3'
        )

    text = sys.argv[1].strip()
    output = Path(sys.argv[2])

    if not text:
        raise SystemExit("TTS text is empty.")

    asyncio.run(synthesize(text, output))

if __name__ == "__main__":
    main()
