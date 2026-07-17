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
import re
import sys
import threading
import time
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


# Cach xac dinh o dang nhap DOffice, THU LAN LUOT tu on dinh nhat den du phong:
#   1) get_by_role + ten hien thi (nhan dang theo ARIA, ghi lai duoc bang
#      Playwright codegen tren trang that: "Tên đăng nhập" / "Mật khẩu" /
#      "Đăng nhập") - CACH KHUYEN NGHI cua Playwright, on dinh voi thay doi
#      giao dien vi dua vao nhan hien thi cho nguoi dung chu khong phai ID noi
#      bo Angular tu sinh.
#   2) ID Angular Material THUC TE da xac nhan tren DOffice (mat-input-0/1,
#      #login-form-wrapper > form > button) - du phong neu nhan hien thi doi
#      ngon ngu/chu.
#   3) Cac selector CSS pho bien (autocomplete, type, name/id chua "user"...)
#      - du phong cuoi cung cho website khac DOffice hoac DOffice doi han giao
#      dien trong tuong lai.
def _username_locator_candidates(page):
    return [
        lambda: page.get_by_role("textbox", name="Tên đăng nhập"),
        lambda: page.locator("#mat-input-0"),
        lambda: page.get_by_role("textbox", name=re.compile("tài khoản|tên đăng nhập|username", re.I)),
        lambda: page.locator('input[autocomplete="username"]'),
        lambda: page.locator('input[type="email"]'),
        lambda: page.locator('input[name*="user" i]'),
        lambda: page.locator('input[id*="user" i]'),
        lambda: page.locator('input[name*="account" i]'),
        lambda: page.locator('input[id*="account" i]'),
        lambda: page.locator('input[placeholder*="tài khoản" i]'),
        lambda: page.locator('input[placeholder*="tên đăng nhập" i]'),
        lambda: page.locator('input[placeholder*="username" i]'),
        lambda: page.locator('input[type="text"]'),
        # Fallback cuoi: bat ky input nao KHONG PHAI password/hidden/checkbox/
        # radio/submit/button - bat truong hop input khong set type="text" tuong
        # minh (Angular Material thuong lam vay).
        lambda: page.locator(
            'input:not([type="password"]):not([type="hidden"]):not([type="checkbox"])'
            ':not([type="radio"]):not([type="submit"]):not([type="button"])'
        ),
    ]


def _password_locator_candidates(page):
    return [
        lambda: page.get_by_role("textbox", name="Mật khẩu"),
        lambda: page.locator("#mat-input-1"),
        lambda: page.locator('input[type="password"]'),
    ]


def _submit_locator_candidates(page):
    return [
        lambda: page.get_by_role("button", name="Đăng nhập"),
        lambda: page.locator("#login-form-wrapper > form > button"),
        lambda: page.locator('button[type="submit"]'),
        lambda: page.locator('input[type="submit"]'),
    ]


def _first_visible(candidates):
    """Thu lan luot 1 danh sach ham tao locator (moi ham 0 tham so, tra ve 1
    Locator), tra ve locator DAU TIEN VISIBLE tim duoc, hoac None neu khong co
    cai nao khop. Dung ham (khong phai chuoi selector) de gom duoc ca
    get_by_role() lan page.locator(css) trong cung 1 danh sach uu tien."""
    for make_locator in candidates:
        try:
            loc = make_locator().first
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return None


def _wait_for_any(candidates, timeout_ms: int):
    """Nhu _first_visible() nhung THU LAI nhieu lan trong toi da timeout_ms (vi
    trang co the con dang render/chuyen huong luc moi vao) - tra ve locator dau
    tien VISIBLE tim duoc, hoac None neu het thoi gian van khong thay."""
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        found = _first_visible(candidates)
        if found is not None:
            return found
        if time.monotonic() >= deadline:
            return None
        time.sleep(0.3)


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
            password_field = _wait_for_any(_password_locator_candidates(page), timeout_ms=10000)
            if password_field is None:
                raise RuntimeError(
                    "Không tìm thấy ô mật khẩu trên trang DOffice trong 10 giây - có thể cấu trúc "
                    "trang khác so với dự kiến. Hãy đăng nhập thủ công trên máy Windows rồi copy "
                    "state.json (xem README_NAS.md)."
                )

            username_field = _first_visible(_username_locator_candidates(page))
            if username_field is None:
                raise RuntimeError(
                    "Tìm thấy ô mật khẩu nhưng không tìm thấy ô tài khoản trên trang DOffice. "
                    "Hãy đăng nhập thủ công trên máy Windows rồi copy state.json (xem README_NAS.md)."
                )

            on_log("Đang điền thông tin đăng nhập...")
            username_field.fill(username)
            password_field.fill(password)

            on_log("Đang bấm đăng nhập...")
            submit_btn = _first_visible(_submit_locator_candidates(page))
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
