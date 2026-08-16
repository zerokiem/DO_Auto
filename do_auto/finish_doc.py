"""
Bam "Ket thuc nhanh" roi "Luu" de dong van ban/cong viec dang xu ly.

Khong dung cho tac vu Van ban Chu tri - Da xu ly (van ban da xong roi, khong can
ket thuc lai) - runner.py se bo qua module nay khi task.enable_finish = False.
"""
from __future__ import annotations

import re

from . import text_utils


_FINISH_TEXT_RE = re.compile(r"K[eêế]t thúc(?: nhanh)?", re.I)


def _visible_finish_menu_item(page):
    """Tra ve menuitem Ket thuc neu menu ba cham dang mo."""
    return page.get_by_role("menuitem", name=_FINISH_TEXT_RE).last


def open_finish_menu(page) -> bool:
    """Mo menu ba cham cua toolbar de hien lenh Ket thuc (khong thuc hien lenh)."""
    menu_item = _visible_finish_menu_item(page)
    try:
        menu_item.wait_for(state="visible", timeout=700)
        return True
    except Exception:
        pass

    # Tren man hinh Xem de biet, trigger dung la nut ba cham (fa-ellipsis-alt)
    # ngay sau nut "Thong tin" tren #vanban > div > mat-toolbar. ID panel do
    # Angular tao dong nen khong dung #mat-menu-panel-41 copy tu DevTools.
    triggers = [
        page.locator(
            '#vanban > div > mat-toolbar > div.ng-star-inserted > '
            'button.mat-menu-trigger[aria-haspopup="true"]'
        ).last,
        page.locator(
            '#vanban mat-toolbar button.mat-menu-trigger[aria-haspopup="true"]'
            ':has(fa-icon[data-icon*="ellipsis"])'
        ).last,
        page.locator(
            '#vanban mat-toolbar button.mat-menu-trigger[aria-haspopup="true"]'
            ':has(.fa-ellipsis-alt)'
        ).last,
        page.locator("#vanban mat-toolbar button[aria-haspopup='true']")
        .last,
    ]
    for trigger in triggers:
        try:
            trigger.wait_for(state="visible", timeout=3000)
            trigger.click(timeout=5000)
            menu_item.wait_for(state="visible", timeout=3000)
            print("✅ Đã mở menu ba chấm có mục Kết thúc.")
            return True
        except Exception:
            continue
    return False


def click_finish_quick_button(page, debug_prefix: str) -> bool:
    """Click nut Ket thuc nhanh/floating button bang nhieu cach."""
    candidates = []

    # DOffice Xem de biet: phai mo menu ba cham truoc, sau do moi co
    # <button role="menuitem"><span>Kết thúc</span></button>.
    if open_finish_menu(page):
        candidates.extend(
            [
                # Span la locator on dinh nhat tren ca hai bien the DOffice da
                # kiem chung. Uu tien no de tranh cho cac fallback 8 giay.
                (
                    page.locator('[id^="mat-menu-panel-"] > div > button[role="menuitem"] > span')
                    .filter(has_text=_FINISH_TEXT_RE)
                    .last,
                    "chữ Kết thúc trong mat-menu-panel",
                ),
                (_visible_finish_menu_item(page), "role menuitem Kết thúc trong menu ba chấm"),
                (
                    page.locator('[id^="mat-menu-panel-"] > div > button[role="menuitem"]')
                    .filter(has_text=_FINISH_TEXT_RE)
                    .last,
                    "button Kết thúc trong mat-menu-panel",
                ),
            ]
        )

    # Cac giao dien cu co nut Ket thuc nhanh truc tiep, khong nam trong menu.
    try:
        candidates.append(
            (page.get_by_role("button", description=re.compile(r"K[eêế]t thúc nhanh", re.I)), "role button description Kết thúc nhanh")
        )
    except TypeError:
        pass

    candidates.extend(
        [
            (page.get_by_role("button", name=re.compile(r"K[eêế]t thúc nhanh", re.I)), "role button name Kết thúc nhanh"),
            (page.locator(".fab-button-ktn > .mat-focus-indicator").first, "floating .fab-button-ktn"),
            # DOffice co bien the dat lenh trong Angular Material menu. ID
            # mat-menu-panel-41 ma Inspect hien thi la ID dong, nen khong hard-code.
            # Selector tuong ung nut: #mat-menu-panel-41 > div > button.mat-menu-item
            (
                page.locator('[id^="mat-menu-panel-"] > div > button.mat-menu-item')
                .filter(has_text=_FINISH_TEXT_RE)
                .last,
                "nút Kết thúc trong mat-menu-panel",
            ),
            # Selector tuong ung text: #mat-menu-panel-41 > div > button.mat-menu-item > span
            (
                page.locator('[id^="mat-menu-panel-"] > div > button.mat-menu-item > span')
                .filter(has_text=_FINISH_TEXT_RE)
                .last,
                "chữ Kết thúc trong mat-menu-panel",
            ),
        ]
    )

    for locator, desc in candidates:
        try:
            locator.wait_for(state="visible", timeout=8000)
            locator.click(timeout=8000)
            print(f"✅ Đã click nút Kết thúc văn bản bằng {desc}.")
            return True
        except Exception as e:
            print(f"⚠️ Không click được {desc}: {e}")

    text_utils.save_debug(page, debug_prefix, "finish_button_failed")
    return False


def click_save_finish_button(page, debug_prefix: str) -> bool:
    """Click nut Luu trong popup Ket thuc nhanh."""
    text_utils.wait(page, 3500)

    dialog = page.locator("mat-dialog-container, .mat-dialog-container, [role='dialog']").last
    candidates = [
        (dialog.get_by_role("button", name=re.compile(r"Lưu", re.I)), "nút Lưu trong dialog"),
        # dialog-ktnhanh-vbd cua man hinh Xem de biet. ID #mat-dialog-6 la
        # dong, vi vay bat dau tu dialog/container thay vi hard-code ID.
        (
            dialog.locator("dialog-ktnhanh-vbd .mat-dialog-actions > button.mat-primary").last,
            "nút Lưu mat-primary trong dialog-ktnhanh-vbd",
        ),
        (
            page.locator("mat-dialog-container dialog-ktnhanh-vbd .mat-dialog-actions > button.mat-primary").last,
            "nút Lưu fallback dialog-ktnhanh-vbd",
        ),
        (page.get_by_role("button", name=re.compile(r"Lưu", re.I)), "role button có chữ Lưu"),
        (page.locator("button").filter(has_text=re.compile(r"Lưu", re.I)).last, "button filter has_text Lưu"),
    ]

    for locator, desc in candidates:
        try:
            locator.wait_for(state="visible", timeout=12000)
            locator.click(timeout=12000)
            print(f"✅ Đã click Lưu xác nhận kết thúc bằng {desc}.")
            return True
        except Exception as e:
            print(f"⚠️ Không click được {desc}: {e}")

    try:
        page.evaluate(
            """
            () => {
                const buttons = Array.from(document.querySelectorAll('button'));
                const btn = buttons.reverse().find(b => /Lưu/i.test((b.innerText || b.textContent || '').trim()));
                if (!btn) throw new Error('Không tìm thấy button Lưu trong DOM');
                btn.click();
            }
            """
        )
        print("✅ Đã click Lưu xác nhận kết thúc bằng JS fallback.")
        return True
    except Exception as e:
        print("❌ Không click được nút Lưu:", e)
        text_utils.save_debug(page, debug_prefix, "save_finish_failed")
        return False


def finish_current_document(page, enable_finish: bool, ask_confirm_before_finish: bool, debug_prefix: str) -> bool:
    print("\n--- Kết thúc văn bản ---")

    if not enable_finish:
        print("⏸️ enable_finish = False, bỏ qua kết thúc văn bản.")
        return True

    if ask_confirm_before_finish:
        ans = input("Chuẩn bị bấm KẾT THÚC văn bản này. Đồng ý? y/n: ").strip().lower()
        if ans != "y":
            print("⏸️ Người dùng bỏ qua kết thúc văn bản.")
            return False

    if not click_finish_quick_button(page, debug_prefix):
        return False

    try:
        # Popup co the chua cum tu nay o ca tieu de va noi dung. Lay phan tu
        # dau tien de khong bi strict-mode violation khi co nhieu ket qua.
        page.get_by_text(re.compile(r"Kết thúc nhanh", re.I)).first.wait_for(state="visible", timeout=12000)
        print("✅ Đã thấy popup Kết thúc nhanh.")
    except Exception:
        print("⚠️ Chưa bắt được text Kết thúc nhanh, vẫn tiếp tục chờ nút Lưu.")

    if ask_confirm_before_finish:
        ans = input("Chuẩn bị bấm LƯU để xác nhận kết thúc. Đồng ý? y/n: ").strip().lower()
        if ans != "y":
            print("⏸️ Người dùng không xác nhận Lưu.")
            return False

    if not click_save_finish_button(page, debug_prefix):
        return False

    try:
        page.locator("mat-dialog-container, .mat-dialog-container, [role='dialog']").last.wait_for(state="hidden", timeout=15000)
    except Exception:
        pass
    text_utils.wait(page, 3000)
    return True
