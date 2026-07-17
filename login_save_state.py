"""
Dang nhap DOffice thu cong 1 lan va luu phien (storage state) de cac script khac
(run_doffice.py) dung lai, khong phai dang nhap lai moi lan chay.

CACH DUNG
---------
    python login_save_state.py

Quy trinh:
    1) Chromium mo ra, tu dong vao DOFFICE_URL trong config.py.
    2) Dang nhap DOffice thu cong (go tai khoan/mat khau, chon chuc danh neu can).
    3) Quay lai cua so terminal, bam Enter.
    4) File playwright/.auth/state.json (hoac duong dan trong config.AUTH_STATE)
       se duoc tao/ghi de.

LUU Y AN TOAN: KHONG chia se, KHONG dua len GitHub file state.json - day la file
luu phien dang nhap, ai co file nay coi nhu dang nhap duoc DOffice thay ban.
"""
from __future__ import annotations

from do_auto.text_utils import fix_windows_console_encoding

fix_windows_console_encoding()

from playwright.sync_api import sync_playwright

import config


def main() -> None:
    config.AUTH_STATE.parent.mkdir(parents=True, exist_ok=True)

    print("=== Đăng nhập DOffice thủ công để lưu phiên ===")
    print(f"Trang đích: {config.DOFFICE_URL}")
    print(f"File phiên sẽ lưu tại: {config.AUTH_STATE.resolve()}")
    print()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(viewport={"width": 1600, "height": 900})
        page = context.new_page()
        page.goto(config.DOFFICE_URL, wait_until="domcontentloaded")

        input(
            "\nChromium đã mở. Hãy đăng nhập DOffice thủ công (và chọn đúng "
            "chức danh nếu DOffice yêu cầu), sau đó quay lại đây và nhấn Enter...\n"
        )

        context.storage_state(path=str(config.AUTH_STATE))
        print(f"\n✅ Đã lưu phiên đăng nhập vào: {config.AUTH_STATE.resolve()}")

        browser.close()


if __name__ == "__main__":
    main()
