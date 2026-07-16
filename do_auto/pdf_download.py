"""
Cho PDF viewer san sang, bam tai xuong, dat ten file than thien va luu ve dia.

Gop tu ban refactor moi nhat (DO_phoi_hop.py / DO_Dang_Doan_phoi_hop_v3.py, hai
ban nay giong het nhau va da toi uu hon ban dau DO_chu_tri_da_XL_v3.py: tach rieng
click_download_button() de moi lan thu co the doi selector, dung danh sach
timeout tang dan thay vi lap co dinh 3 lan).
"""
from __future__ import annotations

import re
import time
from pathlib import Path
from typing import Dict, List, Optional

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from . import text_utils


def wait_pdf_ready(page) -> bool:
    print("--- Chờ PDF viewer sẵn sàng ---")

    try:
        download_btn = page.get_by_role("button", name=re.compile(r"Tải xuống|Download", re.I))
        download_btn.wait_for(state="visible", timeout=15000)
        text_utils.wait(page, 800)
        print("✅ PDF viewer đã sẵn sàng, thấy nút Tải xuống.")
        return True
    except Exception as e:
        print("⚠️ Chưa thấy nút Tải xuống bằng role:", e)

    try:
        page.locator("#download").wait_for(state="visible", timeout=8000)
        text_utils.wait(page, 800)
        print("✅ PDF viewer đã sẵn sàng, thấy #download.")
        return True
    except Exception as e:
        print("❌ Không thấy nút tải PDF:", e)
        return False


def make_friendly_pdf_name(data: Dict[str, str], suggested: str, stt: int) -> str:
    """
    Ten PDF luu xuong: yymmdd-soVB - ten_file_bo_dau.pdf

    Uu tien dung ten file goi y tu DOffice/PDF viewer. Neu ten goi y qua chung
    chung thi fallback theo So VB + Trich yeu.
    """
    suggested = suggested or ""
    base = Path(suggested).name

    generic = not base or re.match(r"^(download|document|file|van_ban|pdf)(\s*\(\d+\))?\.pdf$", base, re.I)
    if generic:
        so_vb = data.get("so_vb", "")
        trich_yeu = data.get("trich_yeu", "")
        base = f"{so_vb} - {trich_yeu}.pdf" if trich_yeu else f"{so_vb}.pdf"

    if base.lower().endswith(".pdf"):
        stem = base[:-4]
    else:
        stem = base
    clean_name = text_utils.safe_ascii_filename(stem) + ".pdf"
    return text_utils.add_date_prefix_if_needed(clean_name, data.get("so_vb", ""))


def click_download_button(page) -> None:
    """Click nut tai xuong trong PDF viewer, thu nhieu selector on dinh."""
    try:
        btn = page.get_by_role("button", name=re.compile(r"Tải xuống|Download", re.I))
        btn.wait_for(state="visible", timeout=5000)
        btn.click(timeout=5000)
        print("✅ Click nút Tải xuống bằng role.")
        return
    except Exception as e:
        print("⚠️ Role Tải xuống không click được:", e)

    try:
        btn = page.locator("#download").first
        btn.wait_for(state="visible", timeout=5000)
        btn.click(timeout=5000)
        print("✅ Click nút Tải xuống bằng #download.")
        return
    except Exception as e:
        print("⚠️ #download không click được:", e)

    page.evaluate(
        """
        () => {
            const btn = document.querySelector('#download')
                || document.querySelector('viewer-download-controls #download')
                || Array.from(document.querySelectorAll('button')).find(b => /Tải xuống|Download/i.test(b.innerText || b.getAttribute('aria-label') || ''));
            if (!btn) throw new Error('Không tìm thấy nút download trong DOM');
            btn.click();
        }
        """
    )
    print("✅ Click nút Tải xuống bằng JS fallback.")


def download_current_document(
    page,
    data: Dict[str, str],
    stt: int,
    download_dir: Path,
    enable_download_pdf: bool,
    attempt_timeouts_ms: List[int],
    debug_prefix: str,
) -> Optional[Path]:
    if not enable_download_pdf:
        print("⏸️ enable_download_pdf = False, bỏ qua tải PDF.")
        return None

    print("\n--- Tải văn bản PDF ---")
    if not wait_pdf_ready(page):
        text_utils.save_debug(page, debug_prefix, "pdf_not_ready")
        return None

    download_dir.mkdir(parents=True, exist_ok=True)

    total_attempts = len(attempt_timeouts_ms)
    for attempt, timeout_ms in enumerate(attempt_timeouts_ms, start=1):
        print(f"Thử tải lần {attempt}/{total_attempts} | timeout {timeout_ms / 1000:.0f}s...")
        try:
            with page.expect_download(timeout=timeout_ms) as download_info:
                click_download_button(page)

            download = download_info.value
            suggested = download.suggested_filename or f"van_ban_{int(time.time())}.pdf"
            friendly_name = make_friendly_pdf_name(data, suggested, stt)
            target = text_utils.unique_path(download_dir, friendly_name)
            download.save_as(str(target))

            print(f"✅ Đã lưu file: {target}")
            return target

        except PlaywrightTimeoutError:
            print("⚠️ Không bắt được download event ở lần này, chuyển nhanh sang lần kế tiếp.")
            text_utils.wait(page, 700)
        except Exception as e:
            print("⚠️ Lỗi tải file:", e)
            text_utils.wait(page, 1000)

    print("❌ Tải file thất bại sau các lần thử.")
    text_utils.save_debug(page, debug_prefix, "download_failed")
    return None
