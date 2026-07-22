"""
Chuyen lich su chay (trang "Lich su" tren web dashboard) tu SQLite (.sqlite3)
sang JSONL (.jsonl) - chay 1 LAN DUY NHAT khi nang cap len ban dung history.py
kieu JSONL (xem do_auto/history.py: SQLite khong lock tin cay qua CIFS/SMB).

Doc toan bo file .sqlite3 cu, ghi lai thanh file .jsonl moi (moi dong 1 JSON,
giu nguyen thu tu/du lieu). File .sqlite3 cu KHONG bi xoa (giu lam ban luu).

An toan chay lai nhieu lan: moi lan chay ghi DE file .jsonl tu dau (khong bi
nhan doi du liệu).

Chay: python migrate_history_to_jsonl.py
"""
from __future__ import annotations

import json
import sqlite3
from pathlib import Path

import config

_COLUMNS = [
    "task_key",
    "task_label",
    "started_at",
    "finished_at",
    "processed",
    "ok",
    "note",
    "trigger_source",
    "log_file",
]


def migrate(sqlite_path: Path, jsonl_path: Path) -> int:
    if not sqlite_path.exists():
        print(f"Không thấy file SQLite cũ: {sqlite_path} - bỏ qua, không có gì để chuyển.")
        return 0

    conn = sqlite3.connect(str(sqlite_path))
    try:
        cur = conn.execute(
            f"SELECT {', '.join(_COLUMNS)} FROM runs ORDER BY id ASC"
        )
        rows = cur.fetchall()
    finally:
        conn.close()

    jsonl_path.parent.mkdir(parents=True, exist_ok=True)
    with open(jsonl_path, "w", encoding="utf-8") as f:
        for row in rows:
            entry = dict(zip(_COLUMNS, row))
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")

    print(f"Đã chuyển {len(rows)} dòng lịch sử từ {sqlite_path.name} sang {jsonl_path.name}.")
    return len(rows)


if __name__ == "__main__":
    base_dir = Path(config.DOWNLOAD_BASE_DIR)
    sqlite_path = base_dir / "doffice_auto_history.sqlite3"
    jsonl_path = base_dir / "doffice_auto_history.jsonl"
    migrate(sqlite_path, jsonl_path)
