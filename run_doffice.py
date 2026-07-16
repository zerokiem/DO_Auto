"""
DOffice Auto - diem vao chinh.

Chon 1, 2 hoac ca 3 tac vu de chay (Van ban Chu tri - Da xu ly / Van ban Phoi hop
- Cho xu ly / Cong viec Dang-Doan Phoi hop - Cho thuc hien). Chi mo/dang nhap
trinh duyet 1 lan cho du chay bao nhieu tac vu.

Doc cau hinh truc tiep tu config.py - sua tay file nay hay sua qua trang web
"Cai dat" deu la CUNG 1 file, khong bi lech giua CLI va web (xem
do_auto/settings_store.py).

CACH DUNG
---------
Menu chon tac vu (khi khong tuyen tham so gi):
    python run_doffice.py

Chon truoc tu dong (dung cho Task Scheduler / chay nhanh khong hoi):
    python run_doffice.py --all
    python run_doffice.py --tasks chu_tri,phoi_hop
    python run_doffice.py --tasks dang_doan

Che do test an toan (ghi de MAX_DOCUMENTS=1, khong bam Ket thuc, hoi xac nhan
truoc moi buoc quan trong - xem config.TEST_MODE_OVERRIDES):
    python run_doffice.py --all --test

Liet ke cac tac vu hien co:
    python run_doffice.py --list
"""
from __future__ import annotations

import argparse
import sys

import config
from do_auto import settings_store
from do_auto.runner import run_selected_tasks


def list_tasks(effective_cfg) -> None:
    print("Các tác vụ hiện có:\n")
    for key, task in sorted(effective_cfg.TASKS.items(), key=lambda kv: kv[1].sheet_order):
        state = "bật" if task.enabled else "TẮT"
        print(f"  {key:<10} - {task.label}  [{state}]")
    print()


def interactive_menu(effective_cfg) -> list[str]:
    ordered = sorted(effective_cfg.TASKS.items(), key=lambda kv: kv[1].sheet_order)
    print("=== DOffice Auto - Chọn công việc cần xử lý ===")
    for i, (key, task) in enumerate(ordered, start=1):
        flag = "" if task.enabled else "  (đang tắt)"
        print(f"  {i}. {task.label}{flag}")
    all_index = len(ordered) + 1
    print(f"  {all_index}. Tất cả ({' + '.join(str(i) for i in range(1, all_index))})")
    print("  0. Thoát")

    raw = input("\nNhập số (có thể chọn nhiều, cách nhau bằng dấu phẩy, ví dụ 1,3): ").strip()
    if not raw or raw == "0":
        print("Đã thoát, không chạy gì cả.")
        sys.exit(0)

    chosen_indexes = set()
    for part in raw.split(","):
        part = part.strip()
        if not part.isdigit():
            continue
        chosen_indexes.add(int(part))

    if all_index in chosen_indexes:
        return [key for key, _ in ordered]

    keys = []
    for idx in sorted(chosen_indexes):
        if 1 <= idx <= len(ordered):
            keys.append(ordered[idx - 1][0])

    if not keys:
        print("Không chọn được tác vụ hợp lệ nào. Đã thoát.")
        sys.exit(0)

    return keys


def apply_test_mode(effective_cfg) -> None:
    print("🧪 CHẾ ĐỘ TEST AN TOÀN: áp dụng config.TEST_MODE_OVERRIDES cho mọi tác vụ.\n")
    effective_cfg.PAUSE_BEFORE_CLOSE = config.TEST_MODE_PAUSE_BEFORE_CLOSE
    effective_cfg.SLOW_MO_MS = config.TEST_MODE_SLOW_MO_MS

    for task in effective_cfg.TASKS.values():
        for field, value in config.TEST_MODE_OVERRIDES.items():
            setattr(task, field, value)


def main() -> None:
    parser = argparse.ArgumentParser(description="DOffice Auto - tải & tổng hợp văn bản DOffice.")
    parser.add_argument(
        "--tasks",
        type=str,
        default=None,
        help="Danh sách tác vụ cách nhau bằng dấu phẩy, ví dụ: chu_tri,phoi_hop,dang_doan",
    )
    parser.add_argument("--all", action="store_true", help="Chạy tất cả tác vụ đang bật")
    parser.add_argument("--list", action="store_true", help="Liệt kê tác vụ hiện có rồi thoát")
    parser.add_argument("--test", action="store_true", help="Chế độ test an toàn (xem config.TEST_MODE_OVERRIDES)")
    parser.add_argument("--no-pause", action="store_true", help="Ép PAUSE_BEFORE_CLOSE=False (đóng browser ngay khi xong)")
    parser.add_argument(
        "--source",
        type=str,
        default="cli",
        choices=["cli", "scheduler"],
        help="Ghi vào lịch sử chạy: 'cli' (chạy tay, mặc định) hay 'scheduler' (Task Scheduler/run_all_doffice.ps1)",
    )
    args = parser.parse_args()

    effective_cfg = settings_store.build_effective_config(config)

    if args.list:
        list_tasks(effective_cfg)
        return

    if args.test:
        apply_test_mode(effective_cfg)

    if args.no_pause:
        effective_cfg.PAUSE_BEFORE_CLOSE = False

    if args.all:
        task_keys = [key for key, task in effective_cfg.TASKS.items() if task.enabled]
    elif args.tasks:
        task_keys = [t.strip() for t in args.tasks.split(",") if t.strip()]
        unknown = [t for t in task_keys if t not in effective_cfg.TASKS]
        if unknown:
            print(f"⚠️ Không nhận diện được tác vụ: {', '.join(unknown)}")
            list_tasks(effective_cfg)
            sys.exit(1)
    else:
        task_keys = interactive_menu(effective_cfg)

    print("\nSẽ chạy các tác vụ:", ", ".join(task_keys))
    print("DOWNLOAD_BASE_DIR:", effective_cfg.DOWNLOAD_BASE_DIR)
    print("EXCEL_FILE:", effective_cfg.EXCEL_FILE)
    print("AUTH_STATE:", effective_cfg.AUTH_STATE)
    print()

    run_selected_tasks(task_keys, effective_cfg, headless=False, trigger_source=args.source)


if __name__ == "__main__":
    main()
