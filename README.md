# Learn Verse — Real Excel Formula Shorts v2

Creates 3 professional 10–15 second vertical Excel Shorts per day.

## What is different from v1?

- Creates a genuine `.xlsx` workbook with OpenPyXL.
- Uses LibreOffice Calc on the GitHub Linux runner to open the real `.xlsx` and show the actual spreadsheet UI.
- Uses Xvfb + xdotool to automate visible data/formula entry while the real spreadsheet window is on screen.
- Records the spreadsheet window with FFmpeg.
- Generates natural-sounding narration with Edge TTS.
- Adds subtitles, hook, result card, sound/transition layers, and Learn Verse branding.
- Produces 1080x1920 MP4 files.
- Uploads the three videos as a GitHub artifact for 3 days.
- Generates a real `.xlsx` beside each video.

## Important limitation

GitHub-hosted Ubuntu runners cannot run Microsoft Excel desktop itself. This project therefore uses **LibreOffice Calc**, which opens and edits the genuine `.xlsx` file format.

The workbook is a real Excel-compatible `.xlsx` file; the visible application is LibreOffice Calc.

If you require the actual Microsoft Excel desktop UI, use a Windows self-hosted GitHub runner with Microsoft Excel installed. Do not add Microsoft Office to a normal Ubuntu GitHub-hosted runner.

## Repository

```text
learnverse-excel-shorts/
├── .github/workflows/generate_excel_shorts.yml
├── assets/learnverse_logo.png
├── output/
├── formula_library.py
├── make_workbook.py
├── record_calc.py
├── make_voice.py
├── render_video.py
├── generate_daily.py
├── requirements.txt
└── README.md
```

## GitHub setup

1. Create a repository, e.g. `LearnVerse-Excel-Shorts`.
2. Upload all files from this project.
3. Make sure the workflow is exactly:
   `.github/workflows/generate_excel_shorts.yml`
4. Commit the files.
5. Open **Actions**.
6. Select **Daily Learn Verse Excel Shorts**.
7. Click **Run workflow** for the first test.
8. Download the `learnverse-excel-shorts` artifact after the run.

The scheduled run is 04:00 UTC, which is 08:00 UAE during UAE standard time (UTC+4).

## Voice

The workflow uses `edge-tts` and does not require a paid API key.

The selected voice is a natural English male neural voice, with narration explaining the problem, formula and result. Change `VOICE` in `make_voice.py` if you prefer another available Edge voice.

Because this is an online TTS service rather than a bundled offline voice, a future service change could require updating the TTS implementation.

## Daily rotation

`formula_library.py` contains workplace examples. `get_daily_formulas()` selects 3 formulas deterministically by date.

Add more formulas to increase the rotation pool.

## Output

Each run creates:

```text
excel_YYYY-MM-DD_01_<formula>.mp4
excel_YYYY-MM-DD_01_<formula>.xlsx
excel_YYYY-MM-DD_02_<formula>.mp4
excel_YYYY-MM-DD_02_<formula>.xlsx
excel_YYYY-MM-DD_03_<formula>.mp4
excel_YYYY-MM-DD_03_<formula>.xlsx
manifest.json
```

The `.xlsx` files are included so you can actually download and use the example spreadsheets.

## Local testing

Ubuntu/Debian:

```bash
sudo apt-get update
sudo apt-get install -y ffmpeg libreoffice xvfb xdotool imagemagick fonts-dejavu
python -m pip install -r requirements.txt
python generate_daily.py --count 1
```

Windows local testing is not the target for the recorder in this version. Use the GitHub Action for the most consistent result.

## Safety / content

All formulas and example results are hard-coded and verified in the library. The video says what the formula is doing and shows the resulting value.
