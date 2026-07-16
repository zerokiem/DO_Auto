"""
Lop "override" nam tren config.py: cho phep sua mot so gia tri thuong dung
(vai tro, so luong van ban toi da, bat/tat tac vu...) tu trang web Cai dat, luu
vao user_settings.json, KHONG can sua code.

config.py van la nguon gia tri MAC DINH (dung khi user_settings.json chua co
hoac chua ghi de truong do). Ca CLI (run_doffice.py) va web dashboard deu goi
build_effective_config() de lay ra bo cau hinh "dang co hieu luc" giong nhau,
nen sua trong Cai dat tren web se ap dung cho ca 2 noi dung, khong bi lech.
"""
from __future__ import annotations

import copy
import json
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict

SETTINGS_FILE = Path(__file__).resolve().parent.parent / "user_settings.json"

# Cac truong cua TaskConfig duoc phep chinh tu web. Co tinh KHONG dua vao danh
# sach nay cac truong dieu huong menu (sidebar_item, list_link, tab_name...) vi
# sua sai co the lam hong hoan toan tac vu - nhung truong do van phai sua trong
# config.py.
TASK_EDITABLE_FIELDS = (
    "enabled",
    "role_pattern",
    "max_documents",
    "enable_finish",
    "ask_confirm_before_finish",
    "enable_download_pdf",
)

# Cac cau hinh chung duoc phep chinh tu web, dang {json_key: ten_thuoc_tinh_tren_cfg}.
COMMON_EDITABLE_FIELDS = {
    "stop_when_duplicate_found": "STOP_WHEN_DUPLICATE_FOUND",
    "duplicate_check_mode": "DUPLICATE_CHECK_MODE",
    "slow_mo_ms": "SLOW_MO_MS",
}


def load_overrides() -> Dict[str, Any]:
    if not SETTINGS_FILE.exists():
        return {"common": {}, "tasks": {}}
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {"common": {}, "tasks": {}}
    data.setdefault("common", {})
    data.setdefault("tasks", {})
    return data


def save_overrides(overrides: Dict[str, Any]) -> None:
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(overrides, f, ensure_ascii=False, indent=2)


def effective_tasks(base_tasks: Dict[str, Any]) -> Dict[str, Any]:
    """Tra ve BAN SAO cua base_tasks (TaskConfig tu config.py), ap dung override
    tu user_settings.json neu co. Khong bao gio sua base_tasks goc."""
    overrides = load_overrides().get("tasks", {})

    result = {}
    for key, task in base_tasks.items():
        task_copy = copy.deepcopy(task)
        for field, value in overrides.get(key, {}).items():
            if field in TASK_EDITABLE_FIELDS and hasattr(task_copy, field):
                setattr(task_copy, field, value)
        result[key] = task_copy
    return result


def effective_common(base_cfg) -> Dict[str, Any]:
    """Tra ve dict cac gia tri chung (khoa = ten thuoc tinh tren cfg, vd
    'SLOW_MO_MS') sau khi ap dung override, KHONG sua doi base_cfg."""
    overrides = load_overrides().get("common", {})
    result = {attr: getattr(base_cfg, attr) for attr in COMMON_EDITABLE_FIELDS.values()}
    for json_key, attr in COMMON_EDITABLE_FIELDS.items():
        if json_key in overrides:
            result[attr] = overrides[json_key]
    return result


def build_effective_config(base_cfg) -> SimpleNamespace:
    """1 doi tuong gom du moi thuoc tinh runner.py can (giong het "shape" cua
    module config.py), da ap dung moi override tu user_settings.json. Dung cho
    CA CLI lan web dashboard de dam bao 2 noi luon thay 1 bo cau hinh nhu nhau."""
    common = effective_common(base_cfg)
    return SimpleNamespace(
        DOFFICE_URL=base_cfg.DOFFICE_URL,
        AUTH_STATE=base_cfg.AUTH_STATE,
        DOWNLOAD_BASE_DIR=base_cfg.DOWNLOAD_BASE_DIR,
        EXCEL_FILE=base_cfg.EXCEL_FILE,
        HISTORY_DB=base_cfg.HISTORY_DB,
        DUPLICATE_CHECK_MODE=common["DUPLICATE_CHECK_MODE"],
        STOP_WHEN_DUPLICATE_FOUND=common["STOP_WHEN_DUPLICATE_FOUND"],
        REFRESH_LIST_EVERY=base_cfg.REFRESH_LIST_EVERY,
        DOWNLOAD_ATTEMPT_TIMEOUTS_MS=base_cfg.DOWNLOAD_ATTEMPT_TIMEOUTS_MS,
        SLOW_MO_MS=common["SLOW_MO_MS"],
        PAUSE_BEFORE_CLOSE=base_cfg.PAUSE_BEFORE_CLOSE,
        ROLE_BUTTON_NAME_HINT=base_cfg.ROLE_BUTTON_NAME_HINT,
        TASKS=effective_tasks(base_cfg.TASKS),
    )
