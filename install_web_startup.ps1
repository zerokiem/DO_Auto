# Cai Windows Scheduled Task de TU DONG chay web dashboard (run_web.py) moi
# khi dang nhap Windows, chay ngam (khong hien cua so console), de ban co the
# mo trinh duyet vao http://127.0.0.1:8877 bat cu luc nao ma khong can tu tay
# go lenh python run_web.py truoc.
#
# Dung "At log on" (khong phai "At startup") vi cac tac vu tu dong hoa (mo
# Chromium that de dang nhap/quan sat) can 1 phien dang nhap Windows that -
# chay o "At startup" (tai khoan SYSTEM, session 0) se KHONG mo duoc cua so
# trinh duyet nhin thay duoc.
#
# Chay PowerShell voi quyen phu hop (Run as Administrator neu gap loi quyen).

$TaskName = "DOffice Web Dashboard"
$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\DO_Auto"
$PythonwExe = Join-Path $ProjectDir ".venv\Scripts\pythonw.exe"
$Script = Join-Path $ProjectDir "run_web.py"

if (!(Test-Path $PythonwExe)) {
    Write-Host "Khong thay pythonw.exe: $PythonwExe"
    Write-Host "Kiem tra lai da tao virtual environment (.venv) chua (xem README muc 3)."
    exit 1
}
if (!(Test-Path $Script)) {
    Write-Host "Khong thay file: $Script"
    exit 1
}

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Action = New-ScheduledTaskAction -Execute $PythonwExe -Argument "`"$Script`"" -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 0) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "Chay ngam DOffice web dashboard moi khi dang nhap Windows" -Force

Write-Host "Da cai: '$TaskName' se tu chay khi dang nhap Windows."
Write-Host ""
Write-Host "Chay ngay bay gio (khong doi den lan dang nhap sau):"
Write-Host "  Start-ScheduledTask -TaskName `"$TaskName`""
Write-Host ""
Write-Host "Dung web dashboard dang chay (khong co nut Stop rieng vi chay ngam khong console):"
Write-Host "  Get-Process pythonw | Stop-Process"
Write-Host ""
Write-Host "Go bo tu dong chay khi dang nhap:"
Write-Host "  Unregister-ScheduledTask -TaskName `"$TaskName`" -Confirm:`$false"
