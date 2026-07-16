# DOffice Auto - PowerShell runner (v5, unified project)
# Goi python run_doffice.py --all thay vi chay 3 script rieng nhu ban cu.
# Van giu co che chong chay chong (lock file) va ghi log.

$ErrorActionPreference = "Stop"

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

# --- CHINH LAI 2 BIEN NAY CHO DUNG TREN MAY BAN ---
$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\DO_Auto"
$PythonExe = Join-Path $ProjectDir ".venv\Scripts\python.exe"

# Tham so truyen cho run_doffice.py. Mac dinh chay ca 3 tac vu (--all).
# --source scheduler de trang Lich su tren web phan biet duoc lan chay nay la
# tu Task Scheduler, khong phai ai do tu go lenh chay tay.
# Vi du chi chay 2 tac vu: "--tasks", "chu_tri,phoi_hop", "--source", "scheduler"
# Them "--test" khi muon chay o che do test an toan.
$RunnerArgs = @("--all", "--no-pause", "--source", "scheduler")

$LogDir = Join-Path $ProjectDir "scheduler_logs"
$LockFile = Join-Path $ProjectDir ".doffice_auto_running.lock"

function NowText {
    return (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

function Write-Log {
    param([string]$Message)
    $line = "[$(NowText)] $Message"
    Write-Host $line
    Add-Content -Path $Global:LogFile -Value $line -Encoding UTF8
}

function Write-RawLog {
    param(
        [string]$Prefix,
        [string]$Text
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $lines = $Text -split "`r?`n"
    foreach ($line in $lines) {
        if ($line -ne "") {
            $out = "[$Prefix] $line"
            Write-Host $out
            Add-Content -Path $Global:LogFile -Value $out -Encoding UTF8
        }
    }
}

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$Global:LogFile = Join-Path $LogDir ("doffice_run_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".log")

try {
    Write-Log "Start DOffice Auto (unified runner)."
    Write-Log "ProjectDir: $ProjectDir"
    Write-Log "PythonExe: $PythonExe"
    Write-Log "Args: $($RunnerArgs -join ' ')"

    if (Test-Path $LockFile) {
        $lockAgeMinutes = ((Get-Date) - (Get-Item $LockFile).LastWriteTime).TotalMinutes
        if ($lockAgeMinutes -lt 180) {
            Write-Log "Another run seems active. Lock age: $([math]::Round($lockAgeMinutes,1)) minutes. Skip this run."
            exit 0
        }
        else {
            Write-Log "Old lock file found. Remove stale lock."
            Remove-Item $LockFile -Force
        }
    }

    Set-Content -Path $LockFile -Value ("Started: " + (NowText)) -Encoding ASCII

    if (!(Test-Path $PythonExe)) {
        throw "Python exe not found: $PythonExe"
    }

    Set-Location $ProjectDir

    $scriptPath = Join-Path $ProjectDir "run_doffice.py"
    if (!(Test-Path $scriptPath)) {
        throw "Script not found: $scriptPath"
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $PythonExe
    $psi.Arguments = ('"' + $scriptPath + '" ' + ($RunnerArgs -join ' '))
    $psi.WorkingDirectory = $ProjectDir
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $false

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    [void]$p.Start()
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    Write-RawLog "STDOUT" $stdout
    Write-RawLog "STDERR" $stderr

    Write-Log "Finished run_doffice.py | ExitCode=$($p.ExitCode)"

    if ($p.ExitCode -ne 0) {
        Write-Log "Run failed. Full log file: $Global:LogFile"
        exit $p.ExitCode
    }

    Write-Log "Run completed successfully."
    Write-Log "Full log file: $Global:LogFile"
}
catch {
    Write-Log "RUNNER ERROR: $($_.Exception.Message)"
    Write-Log "Full log file: $Global:LogFile"
    exit 1
}
finally {
    if (Test-Path $LockFile) {
        Remove-Item $LockFile -Force
        Write-Log "Lock file removed."
    }
}
