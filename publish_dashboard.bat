@echo off
setlocal
cd /d "%~dp0"
python contract_tracker_publish_dashboard.py
if errorlevel 1 (
  echo.
  echo Dashboard refresh did not publish. Review the message above.
  pause
  exit /b 1
)
echo.
echo GitHub Pages is deploying the update to the same dashboard link.
pause
