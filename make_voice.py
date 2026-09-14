import asyncio
import sys
from pathlib import Path
import edge_tts

VOICE = "en-US-AndrewMultilingualNeural"

async def main(text, output):
    communicate = edge_tts.Communicate(text, VOICE, rate="+8%", volume="+0%")
    await communicate.save(str(output))

if __name__ == "__main__":
    asyncio.run(main(sys.argv[1], Path(sys.argv[2])))
