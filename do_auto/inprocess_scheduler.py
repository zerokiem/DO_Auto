"""
Lich chay tu dong CHO MOI TRUONG LINUX/DOCKER (NAS), thay the cho Windows Task
Scheduler. Khong dung cron/supercronic hay bat ky binary/goi ngoai nao - chi
dung 1 thread nen ngay trong tien trinh Flask dang chay (run_web.py), vi
container tren NAS luon chay 24/7 (restart: unless-stopped) nen khong can 1
tien trinh cron rieng.

Danh sach gio duoc luu vao 1 file text don gian (SCHEDULE_FILE, nam trong thu
muc project da mount ra NAS) de giu nguyen qua cac lan container restart.
"""
from __future__ import annotations

import re
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, List, Tuple

PROJECT_DIR = Path(__file__).resolve().parent.parent
SCHEDULE_FILE = PROJECT_DIR / "schedule_times.txt"

_TIME_PATTERN = re.compile(r"^([01]\d|2[0-3]):([0-5]\d)$")

POLL_INTERVAL_SECONDS = 20


def validate_times(times: List[str]) -> Tuple[bool, str]:
    for t in times:
        if not _TIME_PATTERN.match(t):
            return False, f"Giờ không hợp lệ: '{t}' (định dạng đúng là HH:MM, 00:00-23:59)."
    return True, ""


def get_current_times() -> List[str]:
    if not SCHEDULE_FILE.exists():
        return []
    lines = [ln.strip() for ln in SCHEDULE_FILE.read_text(encoding="utf-8").splitlines() if ln.strip()]
    return sorted(set(lines))


def save_times(times: List[str]) -> Tuple[bool, str]:
    ok, err = validate_times(times)
    if not ok:
        return False, err
    if times:
        SCHEDULE_FILE.write_text("\n".join(sorted(set(times))) + "\n", encoding="utf-8")
    elif SCHEDULE_FILE.exists():
        SCHEDULE_FILE.unlink()
    return True, ""


class SchedulerThread(threading.Thread):
    """Thread nen: moi POLL_INTERVAL_SECONDS giay kiem tra gio hien tai co khop
    voi 1 trong cac gio da luu khong. Neu khop VA chua kich hoat cho dung phut
    do (tranh kich hoat 2 lan trong cung 1 phut do khoang poll < 60s), goi
    run_manager.start() giong het nut "Chay" tren web (headless=True, tu dong
    dung ket qua qua cung 1 co che khoa "chi 1 luot chay tai 1 thoi diem")."""

    def __init__(self, run_manager: Any, cfg_module: Any) -> None:
        super().__init__(daemon=True, name="doffice-scheduler")
        self._run_manager = run_manager
        self._cfg_module = cfg_module
        self._last_fired = ""

    def run(self) -> None:
        from do_auto import settings_store

        while True:
            try:
                now = datetime.now()
                current_hm = now.strftime("%H:%M")
                fire_key = now.strftime("%Y-%m-%d %H:%M")

                if current_hm in get_current_times() and fire_key != self._last_fired:
                    self._last_fired = fire_key
                    effective_cfg = settings_store.build_effective_config(self._cfg_module)
                    task_keys = [key for key, task in effective_cfg.TASKS.items() if task.enabled]
                    if task_keys:
                        started = self._run_manager.start(
                            task_keys, headless=True, test_mode=False, trigger_source="scheduler"
                        )
                        if not started:
                            print(
                                f"[scheduler] {fire_key}: bỏ qua vì đang có 1 lượt chạy khác hoạt động."
                            )
            except Exception as e:
                print(f"[scheduler] Lỗi không mong muốn trong vòng lặp lịch chạy: {e}")

            time.sleep(POLL_INTERVAL_SECONDS)


def start_background_scheduler(run_manager: Any, cfg_module: Any) -> SchedulerThread:
    thread = SchedulerThread(run_manager, cfg_module)
    thread.start()
    return thread
