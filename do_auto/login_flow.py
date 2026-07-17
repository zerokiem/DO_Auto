"""
Dang nhap DOffice, dieu khien duoc tu web dashboard - 2 nhanh tuy moi truong:

  * Windows (hoac may Linux co man hinh that, bien DISPLAY): mo Chromium that
    (khong headless), cho nguoi dung dang nhap thu cong, luu phien khi nhan tin
    hieu "da dang nhap xong" - dung threading.Event thay vi cho go Enter trong
    terminal (nhu login_save_state.py ban CLI), vi tien trinh web khong co
    terminal nao de go vao. Xem run_interactive_login().

  * May khong man hinh (NAS/Docker/Pi): KHONG mo duoc cua so that, nen thay vao
    do nhan username/password tu form tren web, tu dien headless va bam dang
    nhap. Xem run_headless_login().

    LUU Y AN TOAN: mat khau CHI dung 1 LAN NGAY LUC DO de dien vao trang dang
    nhap that, KHONG BAO GIO ghi ra dia/log/config.py, khong xuat hien trong bat
    ky thong bao nao gui ve web. Sau khi dang nhap xong (thanh cong hay that
    bai) bien mat khau ra khoi pham vi ham va bi don rac binh thuong.

Ca 2 nhanh chay trong 1 thread nen rieng (giong cach webapp/run_manager.py chay
tac vu tu dong), nen dung duoc Playwright sync API binh thuong.
"""
from __future__ import annotations

import os
import sys
import threading
from typing import Callable

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

from . import browser_nav


def run_interactive_login(cfg, ready_event: threading.Event, on_log: Callable[[str], None] = print) -> None:
    # Tren NAS/Docker (Linux khong co man hinh), khong the mo cua so Chromium that
    # de dang nhap tay. Bao loi ro rang thay vi de Playwright nem loi kho hieu:
    # dang nhap tren Windows roi copy playwright/.auth/state.json len NAS.
    if sys.platform != "win32" and not os.environ.get("DISPLAY"):
        raise RuntimeError(
            "Máy chủ này không có màn hình (NAS/Docker) nên không mở được cửa sổ đăng nhập. "
            "Hãy chạy 'python login_save_state.py' trên máy Windows để đăng nhập, rồi copy file "
            "playwright/.auth/state.json vào thư mục code trên NAS (xem README_NAS.md mục Đăng nhập)."
        )

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


# Trang dang nhap DOffice dung Angular Material - o input KHONG co thuoc tinh
# type="text" tuong minh (Angular chi de trong, trinh duyet mac dinh hieu la
# text) nen selector CSS "input[type=text]" KHONG khop duoc, phai them fallback
# ":not([type=password])..." o cuoi. #mat-input-0/#mat-input-1 la ID DOffice
# THUC TE (xac nhan tren may that) - Angular Material tu danh so tuan tu cac o
# input tren trang, o dang nhap la 2 o dau tien nen on dinh, nhung van de cac
# selector chung phia sau phong khi DOffice doi giao dien.
_USERNAME_SELECTOR_CANDIDATES = [
    "#mat-input-0",  # DOffice hien tai (Angular Material, o dau tien tren form dang nhap)
    'input[autocomplete="username"]',
    'input[type="email"]',
    'input[name*="user" i]',
    'input[id*="user" i]',
    'input[name*="account" i]',
    'input[id*="account" i]',
    'input[placeholder*="tài khoản" i]',
    'input[placeholder*="tên đăng nhập" i]',
    'input[placeholder*="username" i]',
    'input[type="text"]',
    # Fallback cuoi: bat ky input nao KHONG PHAI password/hidden/checkbox/radio/
    # submit/button - bat truong hop nhu DOffice khong set type="text" tuong minh.
    'input:not([type="password"]):not([type="hidden"]):not([type="checkbox"])'
    ':not([type="radio"]):not([type="submit"]):not([type="button"])',
]

_PASSWORD_SELECTOR_CANDIDATES = [
    "#mat-input-1",  # DOffice hien tai
    'input[type="password"]',
]

_SUBMIT_SELECTOR_CANDIDATES = [
    "#login-form-wrapper > form > button",  # DOffice hien tai
    'button[type="submit"]',
    "input[type=\"submit\"]",
]


def _first_visible(page, selectors):
    """Thu lan luot 1 danh sach selector, tra ve locator dau tien VISIBLE tim
    duoc, hoac None neu khong tim duoc kieu nao."""
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return None


def run_headless_login(
    cfg, username: str, password: str, ready_event: threading.Event, on_log: Callable[[str], None] = print
) -> None:
    """Dang nhap DOffice HOAN TOAN TU DONG (khong can nguoi dung tuong tac) -
    dung cho may khong man hinh (NAS/Docker/Pi). ready_event khong dung de cho
    xac nhan (khac voi run_interactive_login) - chi giu tham so cho dong bo chu
    ky ham voi login_manager, luon set() truoc khi ham ket thuc."""
    auth_state = cfg.AUTH_STATE
    auth_state.parent.mkdir(parents=True, exist_ok=True)

    try:
        on_log(f"Đang mở trang đăng nhập DOffice: {cfg.DOFFICE_URL}")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1600, "height": 900})
            page = context.new_page()
            page.goto(cfg.DOFFICE_URL, wait_until="domcontentloaded")

            browser_nav.click_login_if_needed(page)

            on_log("Đang tìm ô đăng nhập...")
            try:
                password_field = page.locator(", ".join(_PASSWORD_SELECTOR_CANDIDATES)).first
                password_field.wait_for(state="visible", timeout=10000)
            except PlaywrightTimeoutError:
                raise RuntimeError(
                    "Không tìm thấy ô mật khẩu trên trang DOffice trong 10 giây - có thể cấu trúc "
                    "trang khác so với dự kiến. Hãy đăng nhập thủ công trên máy Windows rồi copy "
                    "state.json (xem README_NAS.md)."
                )

            username_field = _first_visible(page, _USERNAME_SELECTOR_CANDIDATES)
            if username_field is None:
                raise RuntimeError(
                    "Tìm thấy ô mật khẩu nhưng không tìm thấy ô tài khoản trên trang DOffice. "
                    "Hãy đăng nhập thủ công trên máy Windows rồi copy state.json (xem README_NAS.md)."
                )

            on_log("Đang điền thông tin đăng nhập...")
            username_field.fill(username)
            password_field.fill(password)

            on_log("Đang bấm đăng nhập...")
            submit_btn = _first_visible(page, _SUBMIT_SELECTOR_CANDIDATES)
            if submit_btn is not None:
                submit_btn.click()
            else:
                password_field.press("Enter")

            on_log("Đang chờ kết quả đăng nhập...")
            try:
                password_field.wait_for(state="hidden", timeout=10000)
            except PlaywrightTimeoutError:
                raise RuntimeError(
                    "Đăng nhập không thành công - trang vẫn đang hiện ô mật khẩu sau khi bấm đăng "
                    "nhập (có thể sai tài khoản/mật khẩu). Kiểm tra lại thông tin và thử lại."
                )

            on_log("Đang lưu phiên đăng nhập...")
            context.storage_state(path=str(auth_state))
            on_log(f"✅ Đã đăng nhập và lưu phiên thành công vào: {auth_state}")

            browser.close()
    finally:
        # username/password ra khoi pham vi ham tai day - khong luu, khong log
        # o dau khac trong toan bo ham nay.
        ready_event.set()
