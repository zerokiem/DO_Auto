"""
Dieu phoi 1 lan chay duoc kich hoat tu web: chay trong thread nen (khong lam
treo Flask), phat (broadcast) tung dong log ra cho trinh duyet xem truc tiep
qua Server-Sent Events, va dam bao AN TOAN cho moi truong web (khong bao gio
goi input() vi khong co terminal nao dinh kem de nguoi dung go y/n).

Chi cho phep 1 lan chay hoat dong tai 1 thoi diem (giong co che lock file cua
run_all_doffice.ps1), tranh 2 lan chay cung mo trinh duyet giay vao nhau.
"""
from __future__ import annotations

import io
import queue
import sys
import threading
from collections import deque
from datetime import datetime
from typing import Any, Dict, List, Optional

from do_auto import runner, settings_store


class LogBroadcaster:
    """Phat 1 dong log toi moi trinh duyet dang mo trang dashboard (moi trinh
    duyet co 1 Queue rieng), dong thoi giu lai lich su gan nhat de trinh duyet
    moi mo trang cung thay duoc log cua lan chay dang dien ra."""

    def __init__(self, keep: int = 2000) -> None:
        self._subscribers: List["queue.Queue[str]"] = []
        self._lock = threading.Lock()
        self.history: "deque[str]" = deque(maxlen=keep)

    def publish(self, line: str) -> None:
        self.history.append(line)
        with self._lock:
            subscribers = list(self._subscribers)
        for q in subscribers:
            try:
                q.put_nowait(line)
            except Exception:
                pass

    def subscribe(self) -> "queue.Queue[str]":
        q: "queue.Queue[str]" = queue.Queue()
        with self._lock:
            self._subscribers.append(q)
        return q

    def unsubscribe(self, q: "queue.Queue[str]") -> None:
        with self._lock:
            if q in self._subscribers:
                self._subscribers.remove(q)


class _TeeStream(io.TextIOBase):
    """Ghi ra stdout that (de van thay log neu chay server tu terminal) VA
    dong thoi phat tung dong hoan chinh cho LogBroadcaster."""

    def __init__(self, broadcaster: LogBroadcaster, original) -> None:
        self._broadcaster = broadcaster
        self._original = original
        self._buf = ""

    def write(self, s: str) -> int:
        self._original.write(s)
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line.strip():
                self._broadcaster.publish(line)
        return len(s)

    def flush(self) -> None:
        self._original.flush()


class RunManager:
    def __init__(self, cfg_module: Any) -> None:
        self.cfg_module = cfg_module
        self.broadcaster = LogBroadcaster()
        self._lock = threading.Lock()
        self.is_running = False
        self.current_tasks: List[str] = []
        self.started_at: Optional[str] = None
        self.last_results: List[Any] = []
        self.last_error: Optional[str] = None

    def status(self) -> Dict[str, Any]:
        return {
            "is_running": self.is_running,
            "current_tasks": self.current_tasks,
            "started_at": self.started_at,
            "last_error": self.last_error,
            "last_results": [
                {"key": r.key, "label": r.label, "ok": r.ok, "processed": r.processed, "note": r.note}
                for r in self.last_results
            ],
        }

    def start(self, task_keys: List[str], headless: bool, test_mode: bool, trigger_source: str = "web") -> bool:
        with self._lock:
            if self.is_running:
                return False
            self.is_running = True
            self.current_tasks = task_keys
            self.started_at = datetime.now().isoformat(timespec="seconds")
            self.last_error = None

        thread = threading.Thread(
            target=self._run, args=(task_keys, headless, test_mode, trigger_source), daemon=True
        )
        thread.start()
        return True

    def _run(self, task_keys: List[str], headless: bool, test_mode: bool, trigger_source: str = "web") -> None:
        original_stdout = sys.stdout
        sys.stdout = _TeeStream(self.broadcaster, original_stdout)
        try:
            mode_desc = []
            mode_desc.append("ẩn (headless)" if headless else "hiện cửa sổ")
            if test_mode:
                mode_desc.append("chế độ test")
            self.broadcaster.publish(f"=== Bắt đầu chạy: {', '.join(task_keys)} [{', '.join(mode_desc)}] ===")

            effective_cfg = settings_store.build_effective_config(self.cfg_module)

            if test_mode:
                default_overrides = {"max_documents": 1, "enable_finish": False, "ask_confirm_before_finish": True}
                overrides = getattr(self.cfg_module, "TEST_MODE_OVERRIDES", default_overrides)
                for task in effective_cfg.TASKS.values():
                    for field, value in overrides.items():
                        setattr(task, field, value)
                effective_cfg.SLOW_MO_MS = getattr(self.cfg_module, "TEST_MODE_SLOW_MO_MS", 800)

            # AN TOAN CHO WEB: web server khong co terminal dinh kem de nguoi dung
            # go y/n hay bam Enter, nen LUON tat xac nhan tung buoc va khong bao
            # gio cho dong browser - bat buoc du dang o che do test hay khong.
            effective_cfg.PAUSE_BEFORE_CLOSE = False
            for task in effective_cfg.TASKS.values():
                task.ask_confirm_before_finish = False

            self.last_results = runner.run_selected_tasks(
                task_keys, effective_cfg, headless=headless, trigger_source=trigger_source, test_mode=test_mode
            )
            self.broadcaster.publish("=== Đã chạy xong. Xem tổng kết ở trên hoặc trang Lịch sử. ===")
        except Exception as e:
            self.last_error = str(e)
            self.broadcaster.publish(f"❌ LỖI: {e}")
        finally:
            sys.stdout = original_stdout
            with self._lock:
                self.is_running = False
                self.current_tasks = []
