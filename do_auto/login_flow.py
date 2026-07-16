"""
Dang nhap DOffice thu cong, dieu khien duoc tu web dashboard: mo Chromium that
(khong headless) tren may dang chay server, cho nguoi dung dang nhap, roi luu
phien khi nhan tin hieu "da dang nhap xong" - dung threading.Event thay vi
cho nguoi dung go Enter trong terminal (nhu login_save_state.py ban CLI), vi
tien trinh web khong co terminal nao de nguoi dung go vao.

Chay trong 1 thread nen rieng (giong cach webapp/run_manager.py chay tac vu tu
dong), nen dung duoc Playwright sync API binh thuong.
"""
from __future__ import annotations

import threading
from typing import Callable

from playwright.sync_api import sync_playwright


def run_interactive_login(cfg, ready_event: threading.Event, on_log: Callable[[str], None] = print) -> None:
    auth_state = cfg.AUTH_STATE
    auth_state.parent.mkdir(parents=True, exist_ok=True)

    on_log(f"Đang mở Chromium để đăng nhập DOffice: {cfg.DOFFICE_URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1600, "height": 900})
        page = context.new_page()
        page.goto(cfg.DOFFICE_URL, wait_until="domcontentloaded")

        on_log(
            "Đã mở trang đăng nhập trên máy chủ. Hãy đăng nhập DOffice thủ công trong "
            "cửa sổ Chromium đó, sau đó quay lại trang web và bấm 'Tôi đã đăng nhập xong'."
        )
        ready_event.wait()

        on_log("Đang lưu phiên đăng nhập...")
        context.storage_state(path=str(auth_state))
        on_log(f"✅ Đã lưu phiên đăng nhập vào: {auth_state}")

        browser.close()
