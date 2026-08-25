@echo off
setlocal
cd /d "%~dp0"

echo === Scheduled Task ===
schtasks /Query /TN "DOffice Web Dashboard" /FO LIST 2>nul
if errorlevel 1 echo Khong tim thay task DOffice Web Dashboard.

echo.
echo === Cong 8877 ===
netstat -ano | findstr :8877
if errorlevel 1 echo Khong co tien trinh dang nghe cong 8877.

echo.
echo === Python va phien dang nhap ===
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -c "import config; print('AUTH_STATE =', config.AUTH_STATE); print('EXISTS =', config.AUTH_STATE.exists())"
) else echo Khong thay .venv\Scripts\python.exe
pause
