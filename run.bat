@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating virtual environment...
  py -m venv .venv
  if errorlevel 1 (
    echo Python 3.10+ not found. Install Python and try again.
    pause
    exit /b 1
  )
)
echo Installing requirements...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo Installation failed. Check internet connection or requirements.
  pause
  exit /b 1
)
echo Starting Employee Data App...
".venv\Scripts\python.exe" app.py
if errorlevel 1 pause
