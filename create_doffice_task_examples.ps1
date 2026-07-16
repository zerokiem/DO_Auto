# Tao 1 Windows Scheduled Task DUY NHAT cho DOffice Auto, gan NHIEU thoi diem
# chay trong ngay len CUNG 1 task (khong tao nhieu task rieng) - dung ten task
# giong het do_auto/scheduler.py dung ("DOffice Auto Schedule"), nen trang web
# "Lich chay" (/scheduler) doc/sua duoc task nay binh thuong.
#
# Day la CACH THU CONG (khong can mo web dashboard). Cach thuan tien hon la mo
# http://127.0.0.1:8877/scheduler va dien gio truc tiep tren web - xem README
# muc 7.6 va 8. Chay PowerShell voi quyen phu hop, tot nhat la Run as Administrator
# neu gap loi quyen truy cap.

$TaskName = "DOffice Auto Schedule"
$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
$Runner = Join-Path $ProjectDir "run_all_doffice.ps1"

# Sua danh sach gio chay tai day (co the them/bot, khong gioi han 2 hay 3 gio).
$Times = @("02:30", "12:45", "18:15")

if (!(Test-Path $Runner)) {
    Write-Host "Runner not found: $Runner"
    Write-Host "Copy run_all_doffice.ps1 to ProjectDir first."
    exit 1
}

Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

$Triggers = @()
foreach ($t in $Times) {
    $Triggers += New-ScheduledTaskTrigger -Daily -At $t
}

$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Runner`"" -WorkingDirectory $ProjectDir
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Triggers -Settings $Settings -Description "DOffice Auto - chay tu dong theo lich (--all)" -Force

Write-Host "Da tao task '$TaskName' voi $($Times.Count) thoi diem: $($Times -join ', ')"
Write-Host "Kiem tra: Get-ScheduledTask -TaskName `"$TaskName`" | Get-ScheduledTaskInfo"
Write-Host "Hoac quan ly truc tiep tu trang web: http://127.0.0.1:8877/scheduler"
