@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [LOI] Chua co .venv. Chay setup_windows_cmd.bat truoc.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "login_save_state.py"
if errorlevel 1 (
  echo [LOI] Khong luu duoc phien dang nhap.
  pause
  exit /b 1
)

echo [OK] Da luu phien dang nhap.
pause
