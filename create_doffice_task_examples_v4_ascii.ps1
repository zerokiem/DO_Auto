# Create example Windows Task Scheduler tasks for DOffice Auto v4 runner.
# Run PowerShell as normal user first. If permission issue, run as Administrator.

$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
$Runner = Join-Path $ProjectDir "run_all_doffice_v4_ascii.ps1"

if (!(Test-Path $Runner)) {
    Write-Host "Runner not found: $Runner"
    Write-Host "Copy run_all_doffice_v4_ascii.ps1 to ProjectDir first."
    exit 1
}

$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Runner`"" -WorkingDirectory $ProjectDir
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)

$Times = @(
    @{Name="DOffice Auto 0230"; Time="02:30"},
    @{Name="DOffice Auto 1245"; Time="12:45"},
    @{Name="DOffice Auto 1815"; Time="18:15"}
)

foreach ($t in $Times) {
    $Trigger = New-ScheduledTaskTrigger -Daily -At $t.Time
    Register-ScheduledTask -TaskName $t.Name -Action $Action -Trigger $Trigger -Settings $Settings -Description "Run DOffice auto scripts sequentially" -Force
    Write-Host "Created/updated task: $($t.Name) at $($t.Time)"
}

Write-Host "Done. Open Task Scheduler > Task Scheduler Library to review tasks."
