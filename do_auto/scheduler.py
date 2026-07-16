"""
Quan ly LICH CHAY TU DONG bang 1 Windows Scheduled Task DUY NHAT, gan NHIEU
trigger (nhieu gio trong ngay) len CUNG 1 task - dung nhu ban ghi chu: khong
can nhieu task rieng, 1 task hoan toan co the co nhieu trigger.

Python khong co API Task Scheduler dung san nen goi PowerShell qua subprocess
(Register-ScheduledTask/Get-ScheduledTask/Unregister-ScheduledTask). CHI CHAY
DUOC TREN WINDOWS - cac ham nay se loi ro rang neu goi tren he dieu hanh khac.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Tuple

TASK_NAME = "DOffice Auto Schedule"

_TIME_PATTERN = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")


def _run_powershell(script: str, timeout: int = 30) -> subprocess.CompletedProcess:
    if sys.platform != "win32":
        raise RuntimeError("Tính năng lịch chạy tự động chỉ hoạt động trên Windows (cần powershell.exe/Task Scheduler).")
    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def get_current_times() -> List[str]:
    """Doc cac gio kich hoat hien tai cua task (neu da tao), dang ['HH:MM', ...],
    sap xep tang dan. Tra ve [] neu chua co task nao hoac khong doc duoc."""
    script = f"""
$ErrorActionPreference = 'SilentlyContinue'
$task = Get-ScheduledTask -TaskName '{TASK_NAME}' -ErrorAction SilentlyContinue
if ($null -eq $task) {{ Write-Output '[]'; exit }}
$times = @()
foreach ($trig in $task.Triggers) {{
    if ($trig.StartBoundary) {{
        try {{
            $times += ([datetime]$trig.StartBoundary).ToString('HH:mm')
        }} catch {{}}
    }}
}}
$times | Sort-Object | ConvertTo-Json -Compress
"""
    try:
        result = _run_powershell(script)
    except Exception:
        return []

    raw = (result.stdout or "").strip()
    if not raw:
        return []
    try:
        data = json.loads(raw)
    except Exception:
        return []
    if isinstance(data, str):
        return [data]
    return list(data) if data else []


def apply_schedule(project_dir: Path, times: List[str]) -> Tuple[bool, str]:
    """Xoa task cu (neu co) roi tao lai VOI 1 TASK DUY NHAT co bay nhieu trigger
    tuong ung voi danh sach 'times' (vd ['02:30', '12:45']). Moi trigger chay
    run_all_doffice.ps1 (--all, giong Task Scheduler mau tu truoc)."""
    for t in times:
        if not _TIME_PATTERN.match(t):
            return False, f"Giờ không hợp lệ: '{t}' (định dạng đúng là HH:MM, 00:00-23:59)."

    runner_ps1 = project_dir / "run_all_doffice.ps1"
    if not runner_ps1.exists():
        return False, f"Không thấy file: {runner_ps1}"

    trigger_assignments = "\n".join(f"$t{i} = New-ScheduledTaskTrigger -Daily -At '{t}'" for i, t in enumerate(times))
    trigger_vars = ", ".join(f"$t{i}" for i in range(len(times)))

    script = f"""
$ErrorActionPreference = 'Stop'
Unregister-ScheduledTask -TaskName '{TASK_NAME}' -Confirm:$false -ErrorAction SilentlyContinue

{trigger_assignments}
$Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-NoProfile -ExecutionPolicy Bypass -File "{runner_ps1}"' -WorkingDirectory '{project_dir}'
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 2)
Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $Action -Trigger @({trigger_vars}) -Settings $Settings -Description 'DOffice Auto - chay tu dong theo lich (--all)' -Force | Out-Null
Write-Output 'OK'
"""
    try:
        result = _run_powershell(script, timeout=30)
    except Exception as e:
        return False, str(e)

    ok = result.returncode == 0 and "OK" in (result.stdout or "")
    if ok:
        return True, f"Đã đặt lịch chạy tự động lúc: {', '.join(times)}."
    return False, (result.stderr or result.stdout or "Lỗi không rõ khi tạo lịch.").strip()


def remove_schedule() -> Tuple[bool, str]:
    script = f"""
$ErrorActionPreference = 'SilentlyContinue'
Unregister-ScheduledTask -TaskName '{TASK_NAME}' -Confirm:$false -ErrorAction SilentlyContinue
Write-Output 'OK'
"""
    try:
        result = _run_powershell(script)
    except Exception as e:
        return False, str(e)
    ok = result.returncode == 0
    return ok, "Đã xoá lịch chạy tự động." if ok else (result.stderr or "Lỗi không rõ.").strip()
