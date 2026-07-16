"""
Sua truc tiep vao config.py bang thay the van ban co muc tieu (regex tren tung
truong), giu nguyen moi comment/dinh dang khac trong file - de config.py LUON
la nguon du lieu DUY NHAT. Sua code (mo config.py bang tay) hay sua tren web
(trang Cai dat) deu phan anh dung 1 cho, khong co lop de len tren gay lech giua
2 noi nhu thiet ke ban dau (da bo file user_settings.json).

Chi cho phep sua tu web nhung truong AN TOAN (khong lien quan dieu huong menu -
xem TASK_EDITABLE_FIELDS). Cac truong con lai (sidebar_item, list_link,
tab_name...) van phai mo config.py sua tay, vi sai sot o do co the lam hong han
tac vu.
"""
from __future__ import annotations

import importlib
import json
import re
from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.py"

# Cac truong TaskConfig duoc phep sua tu web, va ham chuyen gia tri Python ->
# van ban se ghi vao config.py.
TASK_FIELD_RENDER = {
    "enabled": lambda v: "True" if v else "False",
    "role_pattern": lambda v: json.dumps(str(v), ensure_ascii=False),
    "max_documents": lambda v: str(int(v)),
    "enable_finish": lambda v: "True" if v else "False",
    "ask_confirm_before_finish": lambda v: "True" if v else "False",
    "enable_download_pdf": lambda v: "True" if v else "False",
}

# Cac bien chung duoc phep sua tu web: {ten_bien_trong_config.py: ham_render}.
COMMON_FIELD_RENDER = {
    "STOP_WHEN_DUPLICATE_FOUND": lambda v: "True" if v else "False",
    "DUPLICATE_CHECK_MODE": lambda v: json.dumps(str(v), ensure_ascii=False),
    "SLOW_MO_MS": lambda v: str(int(v)),
}

_VALUE_PATTERN = r'(?:True|False|"(?:[^"\\]|\\.)*"|-?\d+)'


def _find_task_block(text: str, task_key: str) -> tuple[int, int]:
    """Tra ve (start, end) vi tri khoi TaskConfig(...) ung voi task_key, tinh
    ca dau va cuoi dau ngoac don (can bang dau ngoac, vi trong khoi co the co
    chuoi chua dau ngoac hoac comment)."""
    marker = f'"{task_key}": TaskConfig('
    marker_pos = text.index(marker)
    open_paren = text.index("(", marker_pos)

    depth = 0
    i = open_paren
    while i < len(text):
        ch = text[i]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return marker_pos, i + 1
        i += 1
    raise ValueError(f"Không tìm thấy dấu đóng ngoặc cho tác vụ '{task_key}' trong config.py.")


def update_task_fields(task_key: str, updates: Dict[str, Any]) -> None:
    """Sua 1 hoac nhieu truong cua 1 TaskConfig ngay trong config.py."""
    text = CONFIG_PATH.read_text(encoding="utf-8")
    start, end = _find_task_block(text, task_key)
    block = text[start:end]

    for field, value in updates.items():
        render = TASK_FIELD_RENDER.get(field)
        if render is None:
            continue
        pattern = rf"({re.escape(field)}\s*=\s*){_VALUE_PATTERN}"
        new_block, n = re.subn(pattern, lambda m: m.group(1) + render(value), block, count=1)
        if n == 0:
            raise ValueError(f"Không tìm thấy trường '{field}' trong khối tác vụ '{task_key}' của config.py.")
        block = new_block

    text = text[:start] + block + text[end:]
    CONFIG_PATH.write_text(text, encoding="utf-8")


def update_common_fields(updates: Dict[str, Any]) -> None:
    """Sua 1 hoac nhieu bien cau hinh chung (top-level) ngay trong config.py."""
    text = CONFIG_PATH.read_text(encoding="utf-8")

    for field, value in updates.items():
        render = COMMON_FIELD_RENDER.get(field)
        if render is None:
            continue
        pattern = rf"(^{re.escape(field)}\s*=\s*){_VALUE_PATTERN}"
        new_text, n = re.subn(pattern, lambda m: m.group(1) + render(value), text, count=1, flags=re.MULTILINE)
        if n == 0:
            raise ValueError(f"Không tìm thấy biến '{field}' trong config.py.")
        text = new_text

    CONFIG_PATH.write_text(text, encoding="utf-8")


def reload_config(config_module):
    """Nap lai config.py trong tien trinh Flask dang chay, de thay doi vua ghi
    co hieu luc ngay khong can khoi dong lai server."""
    return importlib.reload(config_module)


def build_effective_config(base_cfg) -> SimpleNamespace:
    """1 doi tuong gom du moi thuoc tinh runner.py can, voi TASKS la BAN SAO
    (deepcopy) cua config.py - de cac dieu chinh chi-danh-cho-1-lan-chay (vd ep
    an toan khi chay qua web trong webapp/run_manager.py) khong bao gio lam
    thay doi config.py that."""
    return SimpleNamespace(
        DOFFICE_URL=base_cfg.DOFFICE_URL,
        AUTH_STATE=base_cfg.AUTH_STATE,
        DOWNLOAD_BASE_DIR=base_cfg.DOWNLOAD_BASE_DIR,
        EXCEL_FILE=base_cfg.EXCEL_FILE,
        HISTORY_DB=base_cfg.HISTORY_DB,
        LOGS_DIR=base_cfg.LOGS_DIR,
        DUPLICATE_CHECK_MODE=base_cfg.DUPLICATE_CHECK_MODE,
        STOP_WHEN_DUPLICATE_FOUND=base_cfg.STOP_WHEN_DUPLICATE_FOUND,
        REFRESH_LIST_EVERY=base_cfg.REFRESH_LIST_EVERY,
        DOWNLOAD_ATTEMPT_TIMEOUTS_MS=base_cfg.DOWNLOAD_ATTEMPT_TIMEOUTS_MS,
        SLOW_MO_MS=base_cfg.SLOW_MO_MS,
        PAUSE_BEFORE_CLOSE=base_cfg.PAUSE_BEFORE_CLOSE,
        ROLE_BUTTON_NAME_HINT=base_cfg.ROLE_BUTTON_NAME_HINT,
        TASKS=deepcopy(base_cfg.TASKS),
    )
