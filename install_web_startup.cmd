@echo off
setlocal EnableExtensions

rem --------------------------------------------------------------
rem Install Windows Scheduled Task to automatically run
rem the DOffice web dashboard when user logs on.
rem --------------------------------------------------------------

rem Get project directory
set "PROJECT_DIR=%~dp0"

rem Remove trailing backslash
if "%PROJECT_DIR:~-1%"=="\" set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"

rem Define task name and paths
set "TaskName=DOffice Web Dashboard"
set "PythonwExe=%PROJECT_DIR%\.venv\Scripts\pythonw.exe"
set "Script=%PROJECT_DIR%\run_web.py"

rem Check pythonw.exe
if not exist "%PythonwExe%" (
    echo.
    echo Khong thay pythonw.exe:
    echo %PythonwExe%
    echo.
    echo Kiem tra lai da tao virtual environment .venv chua.
    pause
    exit /b 1
)

rem Check run_web.py
if not exist "%Script%" (
    echo.
    echo Khong thay file:
    echo %Script%
    pause
    exit /b 1
)

echo.
echo Dang xoa Scheduled Task cu neu co...

schtasks /Delete /TN "%TaskName%" /F >nul 2>&1

echo Dang cai Scheduled Task...

rem Create task at user logon
schtasks /Create ^
 /TN "%TaskName%" ^
 /TR "\"%PythonwExe%\" \"%Script%\"" ^
 /SC ONLOGON ^
 /RL LIMITED ^
 /F

if errorlevel 1 (
    echo.
    echo LOI: Khong tao duoc Scheduled Task.
    echo Neu task cu thuoc tai khoan Administrator, hay xoa no bang CMD Run as administrator.
    echo.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo Da cai thanh cong: %TaskName%
echo Dashboard se tu dong chay khi dang nhap Windows.
echo ==================================================
echo.

echo De chay ngay bay gio, dung lenh:
echo schtasks /Run /TN "%TaskName%"
echo.

pause
endlocal