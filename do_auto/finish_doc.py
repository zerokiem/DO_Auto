"""
Bam "Ket thuc nhanh" roi "Luu" de dong van ban/cong viec dang xu ly.

Khong dung cho tac vu Van ban Chu tri - Da xu ly (van ban da xong roi, khong can
ket thuc lai) - runner.py se bo qua module nay khi task.enable_finish = False.
"""
from __future__ import annotations

import re

from . import text_utils


def click_finish_quick_button(page, debug_prefix: str) -> bool:
    """Click nut Ket thuc nhanh/floating button bang nhieu cach."""
    candidates = []

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
        page.get_by_text(re.compile(r"Kết thúc nhanh", re.I)).wait_for(state="visible", timeout=12000)
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
