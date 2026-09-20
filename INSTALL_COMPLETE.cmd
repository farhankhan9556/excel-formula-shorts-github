@echo off
setlocal
cd /d C:\excel-formula-shorts-github

echo ==========================================
echo LEARN VERSE EXCEL SHORTS FINAL SETUP
echo ==========================================
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m py_compile excel_runner.py record_excel.py excel_actions.py caption_video.py caption_renderer.py make_excel_demo.py formula_library.py make_voice.py youtube_uploader.py
if errorlevel 1 (
  echo.
  echo Python syntax check FAILED.
  pause
  exit /b 1
)
echo.
echo Python syntax check PASSED.
echo.
echo Start the GitHub runner with:
echo cd /d C:\actions-runner
echo .\run.cmd
pause
