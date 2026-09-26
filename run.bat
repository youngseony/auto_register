@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
  echo Setup has not been done yet. Please double-click setup.bat first.
  pause
  exit /b 1
)
"venv\Scripts\python.exe" main.py %*
echo.
pause
