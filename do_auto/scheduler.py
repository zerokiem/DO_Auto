"""
Quan ly LICH CHAY TU DONG (tab "Lich chay" tren web dashboard).

Ho tro 2 moi truong:
  * Windows  -> 1 Windows Scheduled Task DUY NHAT, gan NHIEU trigger (nhieu gio
    trong ngay) len CUNG 1 task (goi PowerShell qua subprocess).
  * Linux/Docker (NAS Synology) -> do_auto/inprocess_scheduler.py: 1 thread nen
    ngay trong tien trinh Flask (khong dung cron/goi binary ngoai nao - container
    tren NAS luon chay 24/7 nen khong can 1 tien trinh cron rieng).

Ca 3 ham cong khai (get_current_times / apply_schedule / remove_schedule) tu
chon dung nhanh theo he dieu hanh, nen webapp/app.py khong can biet dang chay o
dau.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import List, Tuple

from do_auto import inprocess_scheduler

TASK_NAME = "DOffice Auto Schedule"

_TIME_PATTERN = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")


# ==================================================================================
# NHANH WINDOWS (PowerShell / Task Scheduler)
# ==================================================================================
def _run_powershell(script: str, timeout: int = 30) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def _win_get_current_times() -> List[str]:
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


def _win_apply_schedule(project_dir: Path, times: List[str]) -> Tuple[bool, str]:
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


def _win_remove_schedule() -> Tuple[bool, str]:
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


# ==================================================================================
# API CONG KHAI - tu chon nhanh theo he dieu hanh
# ==================================================================================
def get_current_times() -> List[str]:
    if sys.platform == "win32":
        return _win_get_current_times()
    return inprocess_scheduler.get_current_times()


def apply_schedule(project_dir: Path, times: List[str]) -> Tuple[bool, str]:
    for t in times:
        if not _TIME_PATTERN.match(t):
            return False, f"Giờ không hợp lệ: '{t}' (định dạng đúng là HH:MM, 00:00-23:59)."
    if sys.platform == "win32":
        return _win_apply_schedule(project_dir, times)
    ok, err = inprocess_scheduler.save_times(times)
    if not ok:
        return False, err
    return True, f"Đã đặt lịch chạy tự động lúc: {', '.join(times)}."


def remove_schedule() -> Tuple[bool, str]:
    if sys.platform == "win32":
        return _win_remove_schedule()
    ok, err = inprocess_scheduler.save_times([])
    if not ok:
        return False, err
    return True, "Đã xoá lịch chạy tự động."
