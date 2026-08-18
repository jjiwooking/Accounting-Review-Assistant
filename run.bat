@echo off
setlocal
cd /d "%~dp0"

echo [1/3] Python environment check...
if not exist ".venv\Scripts\python.exe" (
  py -m venv .venv 2>nul
  if errorlevel 1 python -m venv .venv
)
if not exist ".venv\Scripts\python.exe" (
  echo Python virtual environment creation failed.
  echo Please install Python 3.11 or later and run again.
  pause
  exit /b 1
)

echo [2/3] Installing/updating required packages...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
if errorlevel 1 (
  echo Package installation failed. Check internet connection and Python installation.
  pause
  exit /b 1
)

echo [3/3] Starting Accounting Review Assistant...
start "" cmd /c "timeout /t 2 >nul & start http://127.0.0.1:8011"
".venv\Scripts\python.exe" -m uvicorn app.main:app --host 127.0.0.1 --port 8011
pause
