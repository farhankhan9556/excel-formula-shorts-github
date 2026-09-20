from __future__ import annotations

import asyncio
from pathlib import Path

import edge_tts

# Adult female neural voices only.
VOICES = [
    "en-US-JennyNeural",
    "en-US-AriaNeural",
    "en-US-SaraNeural",
]

RATE = "+5%"
VOLUME = "+0%"
PITCH = "+1Hz"


async def _save_voice(text: str, voice: str, out: Path) -> None:
    out.unlink(missing_ok=True)
    communicate = edge_tts.Communicate(
        text=text,
        voice=voice,
        rate=RATE,
        volume=VOLUME,
        pitch=PITCH,
    )
    await communicate.save(str(out))


async def make(text, out):
    out = Path(out)
    text = " ".join(str(text).split()).strip()
    if not text:
        raise RuntimeError("TTS text is empty.")

    last_error = None
    for voice in VOICES:
        for attempt in range(3):
            try:
                await _save_voice(text, voice, out)
                if out.exists() and out.stat().st_size > 1000:
                    print(f"VOICE: {voice}")
                    print(f"VOICE RATE: {RATE}")
                    return
            except Exception as exc:
                last_error = exc
                print(f"TTS retry: {voice} attempt {attempt + 1}/3: {exc}")
                await asyncio.sleep(1 + attempt)

    raise RuntimeError(f"TTS failed: {last_error}")
