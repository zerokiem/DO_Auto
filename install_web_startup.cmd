@echo off
rem --------------------------------------------------------------
rem  Install Windows Scheduled Task to automatically run the DOffice web dashboard (run_web.py)
rem  when the user logs on to Windows 10. This batch file provides the same
rem  functionality as install_web_startup.ps1 but works in a plain CMD
rem  environment (no PowerShell required).
rem --------------------------------------------------------------

rem Get the directory of this script (equivalent to $PSScriptRoot)
set "PROJECT_DIR=%~dp0"

rem Define task name and paths
set "TaskName=DOffice Web Dashboard"
set "PythonwExe=%PROJECT_DIR%.venv\Scripts\pythonw.exe"
set "Script=%PROJECT_DIR%run_web.py"

rem Check that pythonw.exe exists
if not exist "%PythonwExe%" (
    echo Khong thay pythonw.exe: %PythonwExe%
    echo Kiem tra lai da tao virtual environment (.venv) chua (xem README muc 3).
    exit /b 1
)

rem Check that run_web.py exists
if not exist "%Script%" (
    echo Khong thay file: %Script%
    exit /b 1
)

rem Remove any existing scheduled task with the same name (ignore errors)
schtasks /Delete /TN "%TaskName%" /F >nul 2>&1

rem Build the command line for the scheduled task. Use double quoting for paths.
set "TaskAction=\"%PythonwExe%\" \"%Script%\""

rem Create the scheduled task to run at user logon. It runs under the current user.
schtasks /Create ^
    /TN "%TaskName%" ^
    /TR "%TaskAction%" ^
    /SC ONLOGON ^
    /RL HIGHEST ^
    /F ^
    /RU "%USERNAME%"

echo Da cai: "%TaskName%" se tu chay khi dang nhap Windows.
echo.
echo Chay ngay bay gio (khong doi den lan dang nhap sau):
echo   schtasks /Run /TN "%TaskName%"
echo.
echo Dung web dashboard dang chay (khong co nut Stop rieng vi chay ngam khong console):
echo   taskkill /IM pythonw.exe /F
echo.
echo Go bo tu dong chay khi dang nhap:
echo   schtasks /Delete /TN "%TaskName%" /F

exit /b 0