# Learn Verse — Excel Reference-Style Shorts V4

This version is designed to reproduce the visual language of the supplied reference:
- vertical 1080x1920 output
- approximately 30 seconds
- Microsoft Excel Desktop visible
- real spreadsheet cells and real Excel Form Controls
- visible mouse/click actions
- timed bottom captions
- AI/TTS narration
- Learn Verse branding
- automatic daily rotation

## Important architecture

Microsoft Excel Desktop is Windows software. GitHub-hosted Ubuntu runners cannot provide the real Excel Desktop interface.

For this V4 setup, GitHub Actions must run on a Windows self-hosted runner/Windows VM that has:
1. Microsoft Excel Desktop installed and activated
2. Python 3
3. FFmpeg in PATH
4. A logged-in Windows desktop session
5. GitHub Actions self-hosted runner installed with labels `self-hosted,windows,excel`

## First test

Run the workflow manually with:
run_date = 2026-09-14
count = 1

Use count=1 for the first test so the Excel UI automation can be verified before running all 3.

## Note about exact content

The system recreates the reference's presentation style and interaction pattern. It does not copy the reference video's exact frames/audio/text.
