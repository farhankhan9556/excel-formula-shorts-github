# Learn Verse Excel Shorts — FINAL ONE-UPLOAD SETUP

This package is the consolidated Learn Verse Excel Shorts setup.

## What it does

- Uses real Microsoft Excel Desktop through COM automation.
- Types real values and formulas into Excel.
- Uses an adult female Microsoft neural voice.
- Records the real Excel desktop with the Windows mouse cursor.
- Final Short is vertical 1080x1920.
- Final Excel view starts at A1 and shows the A:G area.
- Uses three transparent professional information cards:
  - Today's Steps
  - Formula Used
  - Watch on YouTube — Learn Verse / @LearnVerse9556
- Final video duration follows the actual voice instead of forcing 30 seconds.
- Daily formula rotation is deterministic.
- GitHub Actions generates 3 Shorts at 04:00 UTC / 08:00 UAE.
- YouTube uploader uses UAE windows 09:00–10:00, 13:00–14:00, 20:00–21:00.
- Upload state is kept outside the repository in C:\LearnVerseSecrets.

## Windows requirements

- Microsoft Excel Desktop installed.
- FFmpeg available in PATH.
- Python 3.11+.
- GitHub self-hosted runner named `LearnVerse-Excel-PC`.
- Runner labels must include `self-hosted`, `Windows`, `X64`, and `excel`.
- The runner must run in the logged-in interactive Windows desktop because Excel GUI automation and GDI screen capture require the desktop session.

## GitHub upload

Replace the repository contents with the files in this package. Keep `.github/workflows/generate_excel_shorts_windows.yml` in the exact `.github/workflows` path.

Do not upload `client_secret.json`, `token.json`, or any file from `C:\LearnVerseSecrets` to GitHub.

## Manual test

Actions → Learn Verse Excel Shorts → Run workflow → count `1`.

After the 1-video test succeeds, run count `3`.

## Local test

From `C:\excel-formula-shorts-github`:

`python excel_runner.py --count 1`

## YouTube authorization

Put the Google OAuth client file at:

`C:\LearnVerseSecrets\client_secret.json`

Then run:

`python youtube_uploader.py`

The required scope is only:

`https://www.googleapis.com/auth/youtube.upload`

## Important

Keep the self-hosted runner window open while GitHub Actions needs to control Excel.
