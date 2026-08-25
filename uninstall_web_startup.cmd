@echo off
setlocal

set "TaskName=DOffice Web Dashboard"
echo Dang tat va go Scheduled Task: "%TaskName%"
schtasks /End /TN "%TaskName%" >nul 2>&1
schtasks /Delete /TN "%TaskName%" /F >nul 2>&1

rem Chi dung cac PID dang LISTENING o cong 8877 va chi dung neu la Python.
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":8877 .*LISTENING"') do call :stop_python_pid %%P

echo.
echo [OK] Da go task cu va dung web DOffice Python neu dang giu cong 8877.
echo Kiem tra lai bang: netstat -ano ^| findstr :8877
pause
exit /b 0

:stop_python_pid
tasklist /FI "PID eq %1" | findstr /I /R /C:"python\.exe" /C:"pythonw\.exe" >nul
if not errorlevel 1 taskkill /PID %1 /F >nul 2>&1
exit /b 0
