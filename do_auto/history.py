"""
Ghi va doc lich su cac lan chay (dung cho trang "Lich su" tren web dashboard,
nhung khong phu thuoc web - CLI cung ghi lich su nhu nhau qua runner.py).

Dung file .jsonl (moi dong 1 JSON object, chi APPEND, khong bao gio sua/xoa
dong cu) thay vi SQLite: du lieu nay co the dat tren NAS dung chung giua nhieu
may (mount CIFS/SMB) - SQLite can lock (byte-range lock) de doc/ghi an toan,
nhung CIFS khong ho tro lock nay tin cay (hay bao "database is locked" du chi
1 tien trinh dang ghi). Append 1 dong hoan chinh moi lan thi khong can lock gi
ca, an toan qua network filesystem. Xem migrate_history_to_jsonl.py de chuyen
du lieu cu tu ban SQLite truoc do.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import List, Optional


def _parse_line(line: str) -> Optional[dict]:
    line = line.strip()
    if not line:
        return None
    try:
        return json.loads(line)
    except (json.JSONDecodeError, ValueError):
        return None


def record_run(
    db_path: Path,
    task_key: str,
    task_label: str,
    started_at: str,
    finished_at: str,
    processed: int,
    ok: bool,
    note: str = "",
    trigger_source: str = "cli",
    log_file: str = "",
) -> None:
    db_path = Path(db_path)
    db_path.parent.mkdir(parents=True, exist_ok=True)
    entry = {
        "task_key": task_key,
        "task_label": task_label,
        "started_at": started_at,
        "finished_at": finished_at,
        "processed": processed,
        "ok": int(bool(ok)),
        "note": note,
        "trigger_source": trigger_source,
        "log_file": log_file,
    }
    # 1 dong hoan chinh, ket thuc bang \n, mode "a" (O_APPEND) - dam bao khong
    # bi troi/lan voi 1 lan ghi khac dang dien ra cung luc, KHONG can lock nhu
    # SQLite nen an toan qua CIFS/SMB.
    with open(db_path, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def _read_all(db_path: Path) -> List[dict]:
    db_path = Path(db_path)
    if not db_path.exists():
        return []
    entries: List[dict] = []
    with open(db_path, "r", encoding="utf-8") as f:
        for i, line in enumerate(f, start=1):
            entry = _parse_line(line)
            if entry is not None:
                entry.setdefault("id", i)
                entries.append(entry)
    return entries


def list_recent_runs(db_path: Path, limit: int = 100) -> List[dict]:
    entries = _read_all(db_path)
    entries.reverse()
    return entries[:limit]


def list_recent_runs_for_task(db_path: Path, task_key: str, limit: int = 1) -> List[dict]:
    entries = [e for e in _read_all(db_path) if e.get("task_key") == task_key]
    entries.reverse()
    return entries[:limit]
