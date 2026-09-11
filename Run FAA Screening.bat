@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo First-time setup: creating the Python environment...
  python -m venv .venv
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  ".venv\Scripts\python.exe" -m playwright install chromium
)
".venv\Scripts\python.exe" faa_screen.py
echo.
pause
