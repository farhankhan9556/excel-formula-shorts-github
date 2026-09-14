# Learn Verse Excel Shorts — TTS Fix

Replace the existing files with the files in this patch.

## Why the previous run failed

The GitHub log shows the failure occurred in:

`make_voice.py`

before video rendering and artifact upload.

This patch makes TTS much more robust and adds a separate TTS test step. If the TTS service is unavailable, the workflow will stop at the TTS test instead of wasting time generating spreadsheets/videos.

## Files to replace

```text
make_voice.py
requirements.txt
.github/workflows/generate_excel_shorts.yml
```

Keep the rest of your current repository unchanged for this test.

## Important

Keep your logo at:

```text
assets/learnverse_logo.png
```

## Test

After committing the three replacement files:

1. GitHub → Actions
2. Daily Learn Verse Excel Shorts
3. Run workflow
4. Leave `run_date` empty
5. Run workflow

The first step to check is:

`Test Edge TTS before generating videos`

It must show:

`TTS test passed.`

Then the workflow will proceed to generate the 3 videos.
