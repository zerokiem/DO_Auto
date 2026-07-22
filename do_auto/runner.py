"""
Vong lap xu ly chinh cho 1 tac vu (run_task), va dieu phoi chay nhieu tac vu
trong CUNG 1 phien trinh duyet (run_selected_tasks) - chi dang nhap / mo trinh
duyet 1 lan cho du chon 1, 2 hay ca 3 tac vu.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import List

import sys

from playwright.sync_api import sync_playwright

from . import browser_nav, excel_log, extract, finish_doc, history, log_capture, notify, pdf_download, text_utils
from .task_types import TaskConfig, TaskResult


def run_task(page, task: TaskConfig, cfg) -> TaskResult:
    print("\n" + "#" * 70)
    print(f"# TÁC VỤ: {task.label}")
    print("#" * 70)

    download_dir = cfg.DOWNLOAD_BASE_DIR / task.download_subdir
    download_dir.mkdir(parents=True, exist_ok=True)

    excel_log.init_excel_log(cfg.EXCEL_FILE, task.sheet_name, task.title_text)
    existing_keys = excel_log.load_existing_duplicate_keys(cfg.EXCEL_FILE, task.sheet_name, cfg.DUPLICATE_CHECK_MODE)
    existing_filenames = excel_log.load_existing_filenames(cfg.EXCEL_FILE, task.sheet_name)
    print(f"📘 Sheet '{task.sheet_name}' hiện có {len(existing_keys)} khóa văn bản đã tổng hợp.")

    list_status = browser_nav.open_task_list(page, task, cfg)
    if list_status is None:
        print(f"✅ Không còn văn bản cần xử lý cho tác vụ: {task.label}.")
        return TaskResult(task.key, task.label, ok=True, processed=0, note="Không có văn bản cần xử lý")
    if not list_status:
        print(f"❌ Không vào được danh sách cho tác vụ: {task.label}")
        return TaskResult(task.key, task.label, ok=False, processed=0, note="Không mở được danh sách")

    processed = 0
    current_index = 0
    documents: List[dict] = []

    while processed < task.max_documents:
        print("\n" + "=" * 70)
        print(f"[{task.label}] Văn bản thứ {processed + 1} | row {current_index + 1} | đã ghi {processed}/{task.max_documents}")
        print("=" * 70)

        if cfg.REFRESH_LIST_EVERY and processed > 0 and processed % cfg.REFRESH_LIST_EVERY == 0:
            print(f"\n🔄 Đã ghi {processed} văn bản. Load lại danh sách...")
            if not browser_nav.open_task_list(page, task, cfg):
                print("Không load lại được danh sách. Dừng tác vụ này.")
                break
            current_index = 0
            text_utils.wait(page, 2000)

        try:
            row_index_to_pick = 0 if (task.enable_finish and task.always_pick_first_row_after_finish) else current_index
            row = browser_nav.get_document_row(page, row_index_to_pick)
        except Exception as e:
            print("❌ Không lấy được row văn bản:", e)
            print("🔄 Thử load lại danh sách một lần nữa...")
            if not browser_nav.open_task_list(page, task, cfg):
                text_utils.save_debug(page, task.debug_prefix, "cannot_reload_list")
                break
            text_utils.wait(page, 2500)
            try:
                row_index_to_pick = 0 if (task.enable_finish and task.always_pick_first_row_after_finish) else current_index
                row = browser_nav.get_document_row(page, row_index_to_pick)
            except Exception as e2:
                print("❌ Vẫn không lấy được row văn bản sau khi reload:", e2)
                text_utils.save_debug(page, task.debug_prefix, "cannot_get_document_row_after_reload")
                break

        # Thu tu trich xuat / mo van ban khac nhau giua "Da xu ly" (chu_tri) va
        # "Cho xu ly"/"Cho thuc hien" (phoi_hop, dang_doan) - xem ghi chu trong
        # task_types.py o truong check_duplicate_before_open.
        if task.check_duplicate_before_open:
            data = extract.extract_document_info_from_row(row)
            extract.print_document_info(data)

            duplicate_key = excel_log.build_duplicate_key(data, cfg.DUPLICATE_CHECK_MODE)
            if duplicate_key and duplicate_key in existing_keys:
                print(f"🛑 Văn bản đã có trong Excel: {data.get('so_vb', '')}")
                if cfg.STOP_WHEN_DUPLICATE_FOUND:
                    print("Dừng vì đã gặp văn bản trùng (danh sách mới nhất ở trên).")
                    break
                current_index += 1
                continue

            if not browser_nav.click_document_row(row, task.prefer_flag_icon):
                print("Không mở được văn bản. Dừng để kiểm tra.")
                break
        else:
            if not browser_nav.click_document_row(row, task.prefer_flag_icon):
                print("Không mở được văn bản. Dừng để kiểm tra.")
                break

            data = extract.extract_document_info_from_row(row)
            extract.print_document_info(data)

            duplicate_key = excel_log.build_duplicate_key(data, cfg.DUPLICATE_CHECK_MODE)
            if duplicate_key and duplicate_key in existing_keys:
                print(f"🛑 Văn bản đã có trong Excel: {data.get('so_vb', '')}")
                if cfg.STOP_WHEN_DUPLICATE_FOUND:
                    print("Dừng vì đã gặp văn bản trùng (danh sách mới nhất ở trên).")
                    break
                current_index += 1
                continue

        planned_stt = excel_log.get_next_excel_stt(cfg.EXCEL_FILE, task.sheet_name)
        saved_file = pdf_download.download_current_document(
            page,
            data,
            planned_stt,
            download_dir,
            task.enable_download_pdf,
            cfg.DOWNLOAD_ATTEMPT_TIMEOUTS_MS,
            task.debug_prefix,
        )
        if task.enable_download_pdf and not saved_file:
            print("Không tải được văn bản. Dừng để kiểm tra.")
            break

        if task.enable_finish:
            if not finish_doc.finish_current_document(
                page, task.enable_finish, task.ask_confirm_before_finish, task.debug_prefix
            ):
                print("Không kết thúc được văn bản. Dừng để kiểm tra.")
                break

        if saved_file:
            filename_key = text_utils.normalize_key(saved_file.name)
            if filename_key in existing_filenames:
                print(f"🛑 Tên file đã có trong Excel: {saved_file.name}")
                if cfg.STOP_WHEN_DUPLICATE_FOUND:
                    print("Dừng vì gặp tên file trùng.")
                    break
            data["thu_muc_luu"] = excel_log.to_display_folder(
                saved_file.parent, cfg.DOWNLOAD_BASE_DIR, cfg.DISPLAY_BASE_DIR
            )
            data["ten_file_luu"] = saved_file.name
        else:
            data["thu_muc_luu"] = excel_log.to_display_folder(
                download_dir, cfg.DOWNLOAD_BASE_DIR, cfg.DISPLAY_BASE_DIR
            )
            data["ten_file_luu"] = ""

        data["thoi_gian_luu"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        excel_log.append_excel_log(
            cfg.EXCEL_FILE, task.sheet_name, task.title_text, data, getattr(cfg, "DISPLAY_BASE_URL", "")
        )

        documents.append(
            {
                "so_vb": data.get("so_vb", ""),
                "ngay_vb": data.get("ngay_vb", ""),
                "trich_yeu": data.get("trich_yeu", ""),
            }
        )

        existing_keys.add(duplicate_key)
        if data.get("ten_file_luu"):
            existing_filenames.add(text_utils.normalize_key(data["ten_file_luu"]))

        processed += 1
        if task.enable_finish and task.always_pick_first_row_after_finish:
            current_index = 0
        else:
            current_index += 1
        text_utils.wait(page, 1500)

    print(f"\n✅ Hoàn tất tác vụ {task.label}. Đã ghi mới {processed} văn bản.")
    return TaskResult(task.key, task.label, ok=True, processed=processed, documents=documents)


def print_summary(results: List[TaskResult]) -> None:
    print("\n" + "=" * 70)
    print("TỔNG KẾT")
    print("=" * 70)
    for r in results:
        status = "✅ OK" if r.ok else "❌ LỖI"
        extra = f" ({r.note})" if r.note else ""
        print(f"  {status} | {r.label:<55} | ghi mới: {r.processed}{extra}")
    print("=" * 70)


def run_selected_tasks(
    task_keys: List[str],
    cfg,
    *,
    headless: bool = False,
    trigger_source: str = "cli",
    test_mode: bool = False,
) -> List[TaskResult]:
    """Chay 1 hoac nhieu tac vu trong CUNG 1 phien trinh duyet.

    headless: False (mac dinh) = mo cua so Chromium that, giong hanh vi cac
        script goc - phu hop khi chay tay tren may dang ngoi truoc man hinh.
        True = chay an, khong hien cua so - phu hop khi kich hoat tu xa qua web
        dashboard (vd tu dien thoai qua Tailscale), vi luc do khong ai ngoi xem
        cua so Chromium tren may chay server ca.
    trigger_source: ghi vao lich su (HISTORY_DB) cho biet lan chay nay bat nguon
        tu "cli" (chay tay), "scheduler" (Task Scheduler/run_all_doffice.ps1),
        hay "web".
    test_mode: chi de GHI CHU trong tin nhan Telegram tong ket (xem
        do_auto/notify.py) - khong tu no thay doi hanh vi tu dong hoa (cac gia
        tri nhu max_documents=1 phai duoc chinh san trong cfg truoc khi goi ham
        nay, vi du qua run_doffice.py --test).

    MOI LAN GOI HAM NAY (bat ke tu dau) deu tu tao 1 file log rieng trong
    cfg.LOGS_DIR, khong chi rieng khi chay qua Scheduled Task - xem
    do_auto/log_capture.py.
    """
    if not Path(cfg.AUTH_STATE).exists():
        raise FileNotFoundError(
            f"Không thấy file session: {cfg.AUTH_STATE}\n"
            "Chạy: python login_save_state.py   để lưu phiên đăng nhập trước."
        )

    logs_dir = Path(cfg.LOGS_DIR)
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_filename = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{trigger_source}.log"
    log_path = logs_dir / log_filename

    original_stdout = sys.stdout
    log_handle = open(log_path, "a", encoding="utf-8")
    sys.stdout = log_capture.FileTeeStream(original_stdout, log_handle)

    try:
        print(f"=== DOffice Auto | nguồn: {trigger_source} | tác vụ: {', '.join(task_keys)} ===")
        print(f"=== Bắt đầu: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")

        excel_log.ensure_all_sheets(cfg.EXCEL_FILE, cfg.TASKS)

        results: List[TaskResult] = []

        with sync_playwright() as p:
            # Tren Linux/Docker (NAS), Chromium chay bang user root trong container
            # can --no-sandbox; /dev/shm trong container thuong nho nen them
            # --disable-dev-shm-usage de tranh crash tab. Tren Windows khong can.
            launch_args = ["--no-sandbox", "--disable-dev-shm-usage"] if sys.platform != "win32" else []
            browser = p.chromium.launch(headless=headless, slow_mo=cfg.SLOW_MO_MS, args=launch_args)
            context = browser.new_context(
                storage_state=str(cfg.AUTH_STATE),
                accept_downloads=True,
                viewport={"width": 1600, "height": 900},
            )
            page = context.new_page()

            try:
                for key in task_keys:
                    task = cfg.TASKS.get(key)
                    if task is None:
                        print(f"⚠️ Không tìm thấy tác vụ '{key}' trong config.TASKS, bỏ qua.")
                        continue
                    if not task.enabled:
                        print(f"⏸️ Tác vụ '{task.label}' đang enabled=False trong config.py, bỏ qua.")
                        continue

                    started_at = datetime.now()
                    try:
                        result = run_task(page, task, cfg)
                    except Exception as e:
                        print(f"❌ Lỗi không mong muốn khi chạy tác vụ {task.label}: {e}")
                        text_utils.save_debug(page, task.debug_prefix, "unexpected_error")
                        result = TaskResult(task.key, task.label, ok=False, processed=0, note=str(e))
                    finished_at = datetime.now()

                    try:
                        history.record_run(
                            Path(cfg.HISTORY_DB),
                            task_key=result.key,
                            task_label=result.label,
                            started_at=started_at.strftime("%Y-%m-%d %H:%M:%S"),
                            finished_at=finished_at.strftime("%Y-%m-%d %H:%M:%S"),
                            processed=result.processed,
                            ok=result.ok,
                            note=result.note,
                            trigger_source=trigger_source,
                            log_file=log_filename,
                        )
                    except Exception as e:
                        print(f"⚠️ Không ghi được lịch sử chạy: {e}")

                    results.append(result)
            finally:
                if cfg.PAUSE_BEFORE_CLOSE:
                    input("\nNhấn Enter để đóng browser...")
                browser.close()

        print_summary(results)

        try:
            notify.notify_run_summary(cfg, results, trigger_source, test_mode=test_mode)
        except Exception as e:
            print(f"⚠️ Không gửi được thông báo tổng kết: {e}")

        print(f"\n=== Kết thúc: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
        print(f"=== File log này: {log_path} ===")
        return results
    finally:
        sys.stdout = original_stdout
        log_handle.close()
