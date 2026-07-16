"""
Ghi va doc lich su cac lan chay (dung cho trang "Lich su" tren web dashboard,
nhung khong phu thuoc web - CLI cung ghi lich su nhu nhau qua runner.py).

Dung SQLite (1 file .sqlite3, khong can cai dat gi them) vi day la du lieu van
hanh noi bo, doc/ghi it, khong can server database rieng.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import List, Optional


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_key TEXT NOT NULL,
            task_label TEXT NOT NULL,
            started_at TEXT NOT NULL,
            finished_at TEXT NOT NULL,
            processed INTEGER NOT NULL DEFAULT 0,
            ok INTEGER NOT NULL DEFAULT 0,
            note TEXT DEFAULT '',
            trigger_source TEXT DEFAULT 'cli'
        )
        """
    )
    return conn


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
) -> None:
    conn = _connect(db_path)
    try:
        with conn:
            conn.execute(
                """
                INSERT INTO runs
                    (task_key, task_label, started_at, finished_at, processed, ok, note, trigger_source)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (task_key, task_label, started_at, finished_at, processed, int(ok), note, trigger_source),
            )
    finally:
        conn.close()


_COLUMNS = [
    "id",
    "task_key",
    "task_label",
    "started_at",
    "finished_at",
    "processed",
    "ok",
    "note",
    "trigger_source",
]


def list_recent_runs(db_path: Path, limit: int = 100) -> List[dict]:
    if not db_path.exists():
        return []
    conn = _connect(db_path)
    try:
        cur = conn.execute(
            f"SELECT {', '.join(_COLUMNS)} FROM runs ORDER BY id DESC LIMIT ?",
            (limit,),
        )
        return [dict(zip(_COLUMNS, row)) for row in cur.fetchall()]
    finally:
        conn.close()


def list_recent_runs_for_task(db_path: Path, task_key: str, limit: int = 1) -> List[dict]:
    if not db_path.exists():
        return []
    conn = _connect(db_path)
    try:
        cur = conn.execute(
            f"SELECT {', '.join(_COLUMNS)} FROM runs WHERE task_key = ? ORDER BY id DESC LIMIT ?",
            (task_key, limit),
        )
        return [dict(zip(_COLUMNS, row)) for row in cur.fetchall()]
    finally:
        conn.close()
