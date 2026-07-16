"""
Sua truc tiep vao config.py bang thay the van ban co muc tieu (regex tren tung
truong), giu nguyen moi comment/dinh dang khac trong file - de config.py LUON
la nguon du lieu DUY NHAT. Sua code (mo config.py bang tay) hay sua tren web
(trang Cai dat) deu phan anh dung 1 cho, khong co lop de len tren gay lech giua
2 noi (da bo file user_settings.json).

Chi cho phep sua tu web nhung truong AN TOAN (khong lien quan dieu huong menu -
xem TASK_EDITABLE_FIELDS). Cac truong con lai (sidebar_item, list_link,
tab_name...) van phai mo config.py sua tay, vi sai sot o do co the lam hong han
tac vu.

TUONG THICH NGUOC: nguoi dung thuong GIU LAI config.py da tuy chinh cua ho thay
vi ghi de bang ban template moi khi cap nhat code - hop ly, vi ho da dien san
role/duong dan rieng. Nhung dieu do co nghia config.py cua ho co the CHUA CO
nhung bien moi them sau nay (vd ENABLE_TELEGRAM_NOTIFY). De tranh loi
AttributeError lam sap trang web, moi truong doc trong build_effective_config()
deu dung getattr() kem gia tri mac dinh, va ham ghi (update_common_fields /
update_task_fields) se TU DONG THEM dong con thieu vao config.py thay vi bao
loi, ngay lan dau nguoi dung bam Luu tren trang Cai dat.
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
    "DOFFICE_URL": lambda v: json.dumps(str(v), ensure_ascii=False),
    "STOP_WHEN_DUPLICATE_FOUND": lambda v: "True" if v else "False",
    "DUPLICATE_CHECK_MODE": lambda v: json.dumps(str(v), ensure_ascii=False),
    "SLOW_MO_MS": lambda v: str(int(v)),
    "ENABLE_TELEGRAM_NOTIFY": lambda v: "True" if v else "False",
    "TELEGRAM_BOT_TOKEN": lambda v: json.dumps(str(v), ensure_ascii=False),
    "TELEGRAM_CHAT_ID": lambda v: json.dumps(str(v), ensure_ascii=False),
    "TELEGRAM_NOTIFY_ONLY_IF_NEW": lambda v: "True" if v else "False",
}

# Gia tri mac dinh dung khi config.py cua nguoi dung CHUA CO bien nay (ban cu
# hon lan them bien do) - chi de KHONG BI SAP TRANG, khong thay the cho viec
# dien cau hinh that (vd Telegram se coi nhu dang tat cho toi khi nguoi dung
# tu dien qua trang Cai dat).
_DEFAULTS: Dict[str, Any] = {
    "DOFFICE_URL": "https://doffice.npt.com.vn/",
    "ROLE_BUTTON_NAME_HINT": "",
    "REFRESH_LIST_EVERY": 0,
    "DOWNLOAD_ATTEMPT_TIMEOUTS_MS": [7000, 15000, 20000],
    "SLOW_MO_MS": 250,
    "PAUSE_BEFORE_CLOSE": False,
    "STOP_WHEN_DUPLICATE_FOUND": True,
    "DUPLICATE_CHECK_MODE": "so_vb_ngay_vb",
    "ENABLE_TELEGRAM_NOTIFY": False,
    "TELEGRAM_BOT_TOKEN": "",
    "TELEGRAM_CHAT_ID": "",
    "TELEGRAM_NOTIFY_ONLY_IF_NEW": False,
    "DISPLAY_BASE_DIR": None,
    "DISPLAY_BASE_URL": "",
}

_VALUE_PATTERN = r'(?:True|False|"(?:[^"\\]|\\.)*"|-?\d+)'
_AUTO_APPEND_MARKER = "# --- Tự động thêm bởi trang web Cài đặt (biến cấu hình mới từ bản cập nhật) ---"


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
    """Sua 1 hoac nhieu truong cua 1 TaskConfig ngay trong config.py. Truong
    nao chua ton tai trong khoi (config.py cu hon) se duoc TU DONG THEM vao
    ngay truoc dau ')' dong khoi, thay vi bao loi."""
    text = CONFIG_PATH.read_text(encoding="utf-8")
    start, end = _find_task_block(text, task_key)
    block = text[start:end]

    to_append = []
    for field, value in updates.items():
        render = TASK_FIELD_RENDER.get(field)
        if render is None:
            continue
        pattern = rf"({re.escape(field)}\s*=\s*){_VALUE_PATTERN}"
        new_block, n = re.subn(pattern, lambda m: m.group(1) + render(value), block, count=1)
        if n == 0:
            to_append.append(f"        {field}={render(value)},")
            continue
        block = new_block

    if to_append:
        close_idx = block.rfind(")")
        if close_idx == -1:
            raise ValueError(f"Không tìm thấy dấu ')' đóng khối tác vụ '{task_key}' trong config.py.")
        block = block[:close_idx] + "\n".join(to_append) + "\n    " + block[close_idx:]

    text = text[:start] + block + text[end:]
    CONFIG_PATH.write_text(text, encoding="utf-8")


def update_common_fields(updates: Dict[str, Any]) -> None:
    """Sua 1 hoac nhieu bien cau hinh chung (top-level) ngay trong config.py.
    Bien nao chua ton tai (config.py cu hon) se duoc TU DONG THEM vao cuoi
    file, thay vi bao loi va buoc nguoi dung tu sua tay."""
    text = CONFIG_PATH.read_text(encoding="utf-8")

    to_append = []
    for field, value in updates.items():
        render = COMMON_FIELD_RENDER.get(field)
        if render is None:
            continue
        pattern = rf"(^{re.escape(field)}\s*=\s*){_VALUE_PATTERN}"
        new_text, n = re.subn(pattern, lambda m: m.group(1) + render(value), text, count=1, flags=re.MULTILINE)
        if n == 0:
            to_append.append(f"{field} = {render(value)}")
            continue
        text = new_text

    if to_append:
        text = text.rstrip("\n") + "\n\n"
        if _AUTO_APPEND_MARKER not in text:
            text += _AUTO_APPEND_MARKER + "\n"
        text += "\n".join(to_append) + "\n"

    CONFIG_PATH.write_text(text, encoding="utf-8")


def reload_config(config_module):
    """Nap lai config.py trong tien trinh Flask dang chay, de thay doi vua ghi
    co hieu luc ngay khong can khoi dong lai server."""
    return importlib.reload(config_module)


def _get(base_cfg, name: str) -> Any:
    """getattr voi fallback ve _DEFAULTS - khong bao gio nem AttributeError du
    config.py cua nguoi dung co cu hon code hien tai bao nhieu di nua."""
    return getattr(base_cfg, name, _DEFAULTS.get(name))


def build_effective_config(base_cfg) -> SimpleNamespace:
    """1 doi tuong gom du moi thuoc tinh runner.py can, voi TASKS la BAN SAO
    (deepcopy) cua config.py - de cac dieu chinh chi-danh-cho-1-lan-chay (vd ep
    an toan khi chay qua web trong webapp/run_manager.py) khong bao gio lam
    thay doi config.py that.

    Moi truong "tuy chon" (co the chua co trong config.py cu) deu doc qua
    _get() de khong bao gio lam sap trang web vi AttributeError - xem docstring
    dau file."""
    download_base_dir = base_cfg.DOWNLOAD_BASE_DIR

    return SimpleNamespace(
        DOFFICE_URL=_get(base_cfg, "DOFFICE_URL"),
        AUTH_STATE=base_cfg.AUTH_STATE,
        DOWNLOAD_BASE_DIR=download_base_dir,
        # Duong dan hien thi de bam mo file tren Windows (o S:). Config cu chua co
        # bien nay -> fallback ve chinh download_base_dir (giu nguyen hanh vi cu).
        DISPLAY_BASE_DIR=getattr(base_cfg, "DISPLAY_BASE_DIR", None) or str(download_base_dir),
        # URL web toi NAS de bam mo file tren dien thoai/Tailscale. "" -> giu link file://.
        DISPLAY_BASE_URL=getattr(base_cfg, "DISPLAY_BASE_URL", "") or "",
        EXCEL_FILE=base_cfg.EXCEL_FILE,
        HISTORY_DB=getattr(base_cfg, "HISTORY_DB", download_base_dir / "doffice_auto_history.sqlite3"),
        LOGS_DIR=getattr(base_cfg, "LOGS_DIR", download_base_dir / "logs"),
        DUPLICATE_CHECK_MODE=_get(base_cfg, "DUPLICATE_CHECK_MODE"),
        STOP_WHEN_DUPLICATE_FOUND=_get(base_cfg, "STOP_WHEN_DUPLICATE_FOUND"),
        REFRESH_LIST_EVERY=_get(base_cfg, "REFRESH_LIST_EVERY"),
        DOWNLOAD_ATTEMPT_TIMEOUTS_MS=_get(base_cfg, "DOWNLOAD_ATTEMPT_TIMEOUTS_MS"),
        SLOW_MO_MS=_get(base_cfg, "SLOW_MO_MS"),
        PAUSE_BEFORE_CLOSE=_get(base_cfg, "PAUSE_BEFORE_CLOSE"),
        ROLE_BUTTON_NAME_HINT=_get(base_cfg, "ROLE_BUTTON_NAME_HINT"),
        ENABLE_TELEGRAM_NOTIFY=_get(base_cfg, "ENABLE_TELEGRAM_NOTIFY"),
        TELEGRAM_BOT_TOKEN=_get(base_cfg, "TELEGRAM_BOT_TOKEN"),
        TELEGRAM_CHAT_ID=_get(base_cfg, "TELEGRAM_CHAT_ID"),
        TELEGRAM_NOTIFY_ONLY_IF_NEW=_get(base_cfg, "TELEGRAM_NOTIFY_ONLY_IF_NEW"),
        TASKS=deepcopy(base_cfg.TASKS),
    )
