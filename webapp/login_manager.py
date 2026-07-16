"""
Dieu phoi 1 luot dang nhap lai duoc kich hoat tu web dashboard - xem
do_auto/login_flow.py cho phan chay Playwright thuc te.
"""
from __future__ import annotations

import threading
from typing import Any, Dict, Optional

from do_auto import login_flow


class LoginManager:
    def __init__(self, cfg_module: Any) -> None:
        self.cfg_module = cfg_module
        self._lock = threading.Lock()
        self.is_active = False
        self._event: Optional[threading.Event] = None
        self.message = ""
        self.error: Optional[str] = None

    def status(self) -> Dict[str, Any]:
        return {"is_active": self.is_active, "message": self.message, "error": self.error}

    def start(self) -> bool:
        with self._lock:
            if self.is_active:
                return False
            self.is_active = True
            self._event = threading.Event()
            self.message = "Đang mở cửa sổ đăng nhập trên máy chủ..."
            self.error = None

        thread = threading.Thread(target=self._run, daemon=True)
        thread.start()
        return True

    def confirm(self) -> bool:
        with self._lock:
            if not self.is_active or self._event is None:
                return False
            self._event.set()
            return True

    def _run(self) -> None:
        try:
            login_flow.run_interactive_login(self.cfg_module, self._event, on_log=self._log)
            self.message = "Đã lưu phiên đăng nhập thành công."
        except Exception as e:
            self.error = str(e)
            self.message = f"Lỗi khi đăng nhập: {e}"
        finally:
            with self._lock:
                self.is_active = False

    def _log(self, msg: str) -> None:
        self.message = msg
