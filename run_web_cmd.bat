@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [LOI] Chua co .venv. Chay setup_windows_cmd.bat truoc.
  pause
  exit /b 1
)

echo DOffice web dang chay tai http://127.0.0.1:8877
echo Giu cua so nay mo. Bam Ctrl+C de dung.
echo.
".venv\Scripts\python.exe" "run_web.py"
pause
