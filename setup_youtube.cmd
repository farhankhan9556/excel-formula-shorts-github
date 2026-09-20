@echo off
setlocal
cd /d C:\excel-formula-shorts-github
if not exist C:\LearnVerseSecrets mkdir C:\LearnVerseSecrets
python -m pip install -r requirements.txt
echo.
echo ==========================================
echo LEARN VERSE YOUTUBE SETUP
echo ==========================================
echo Put your Google OAuth client_secret.json here:
echo C:\LearnVerseSecrets\client_secret.json
echo.
echo Then run:
echo python youtube_uploader.py
echo.
echo OAuth scope used: youtube.upload
echo ==========================================
pause
