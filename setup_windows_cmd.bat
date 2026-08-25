@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   DOffice Auto - Cai dat Windows 10 (CMD)
echo ============================================
echo.

where py >nul 2>nul
if not errorlevel 1 goto use_py
where python >nul 2>nul
if not errorlevel 1 goto use_python

echo [LOI] Khong tim thay Python.
echo Cai Python 3.11 hoac 3.12, tick Add Python to PATH, roi chay lai.
pause
exit /b 1

:use_py
set "PYTHON_CMD=py"
set "PYTHON_ARG=-3"
goto create_venv

:use_python
set "PYTHON_CMD=python"
set "PYTHON_ARG="

:create_venv
if exist ".venv\Scripts\python.exe" goto install_packages
echo [1/3] Tao moi truong .venv...
"%PYTHON_CMD%" %PYTHON_ARG% -m venv ".venv"
if errorlevel 1 goto failed

:install_packages
echo [2/3] Cai thu vien Python...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto failed
".venv\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto failed

echo [3/3] Cai Chromium cho Playwright...
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 goto failed

if not exist "playwright\.auth" mkdir "playwright\.auth"
echo.
echo [OK] Cai dat xong.
echo Chay login_doffice.bat de tao phien dang nhap.
pause
exit /b 0

:failed
echo.
echo [LOI] Cai dat that bai. Xem thong bao phia tren.
pause
exit /b 1
