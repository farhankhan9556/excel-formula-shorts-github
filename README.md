# Learn Verse Excel Shorts — Stable GitHub Automation

This repository generates 3 Excel Shorts per day.

## Important
Keep the existing logo at:

`assets/learnverse_logo.png`

The workflow uses LibreOffice Calc on a real `.xlsx` workbook and FFmpeg/Xvfb to record the spreadsheet screen.

## Daily schedule
The workflow runs at `04:00 UTC`, which is `08:00 UAE`.

## Manual test
GitHub → Actions → Daily Learn Verse Excel Shorts → Run workflow.

Use:
- `run_date`: `2026-09-14`
- `count`: `3`

## Output
Each successful run creates:
- 3 MP4 Shorts
- 3 Excel workbooks
- `manifest.json`

Artifacts are retained for 3 days.

## Do not delete
Keep:
`assets/learnverse_logo.png`
