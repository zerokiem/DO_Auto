"""
Sua truc tiep vao config.py bang thay the van ban co muc tieu (regex tren tung
truong), giu nguyen moi comment/dinh dang khac trong file - de config.py LUON
la nguon du lieu DUY NHAT. Sua code (mo config.py bang tay) hay sua tren web
(trang Cai dat) deu phan anh dung 1 cho, khong co lop de len tren gay lech giua
2 noi (da bo file user_settings.json).

Trang web cho phep sua luong dieu huong don gian (sidebar_item, list_link,
tab_name) va tuy chon JSON nang cao. Moi gia tri deu duoc validate/ghi qua AST
de selector co ngoac hoac chu tieng Viet khong lam hong config.py.

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

import ast
import importlib
import json
import pprint
import re
from copy import deepcopy
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.py"

# Cac truong TaskConfig duoc phep sua tu web, va ham chuyen gia tri Python ->
# van ban se ghi vao config.py.
TASK_FIELD_RENDER = {
    "label": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "enabled": lambda v: "True" if v else "False",
    "role_pattern": lambda v: json.dumps(str(v), ensure_ascii=False),
    "sidebar_item": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "list_link": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "list_link_href": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "tab_name": lambda v: "None" if v in (None, "") else json.dumps(str(v).strip(), ensure_ascii=False),
    # config.py la Python, khong phai JSON: phai ghi True/False/None dung cu
    # phap Python. json.dumps() se tao true/false/null va lam config.py hong.
    "navigation_steps": lambda v: pprint.pformat(v, width=120, sort_dicts=False),
    "use_advanced_navigation": lambda v: "True" if v else "False",
    "document_row_selector": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "document_click_selector": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "extract_mode": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "max_documents": lambda v: str(int(v)),
    "enable_finish": lambda v: "True" if v else "False",
    "ask_confirm_before_finish": lambda v: "True" if v else "False",
    "enable_download_pdf": lambda v: "True" if v else "False",
    "download_subdir": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "sheet_name": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "title_text": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "sheet_order": lambda v: str(int(v)),
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
    "DISPLAY_BASE_DIR_OVERRIDE": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
    "DISPLAY_BASE_URL_OVERRIDE": lambda v: json.dumps(str(v).strip(), ensure_ascii=False),
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
    "DISPLAY_BASE_DIR_OVERRIDE": "",
    "DISPLAY_BASE_URL_OVERRIDE": "",
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


def _offset(source: str, lineno: int, col_offset: int) -> int:
    """Doi vi tri AST (1-based line, cot UTF-8) sang offset chuoi.

    ast dung byte offset UTF-8, khong phai so ky tu Python. Dieu nay rat quan
    trong vi label/selector cua DOffice thuong co tieng Viet co dau.
    """
    lines = source.splitlines(keepends=True)
    line = lines[lineno - 1]
    char_column = len(line.encode("utf-8")[:col_offset].decode("utf-8"))
    return sum(len(item) for item in lines[: lineno - 1]) + char_column


def _tasks_dict(source: str):
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.Assign) and any(isinstance(t, ast.Name) and t.id == "TASKS" for t in node.targets):
            if isinstance(node.value, ast.Dict):
                return node.value
    raise ValueError("Không tìm thấy TASKS = {...} trong config.py.")


def _task_call(source: str, task_key: str):
    tasks = _tasks_dict(source)
    for key_node, value_node in zip(tasks.keys, tasks.values):
        if isinstance(key_node, ast.Constant) and key_node.value == task_key and isinstance(value_node, ast.Call):
            return key_node, value_node
    raise ValueError(f"Không tìm thấy tác vụ '{task_key}' trong config.py.")


def _validate_task_key(task_key: str) -> str:
    key = str(task_key or "").strip()
    if not re.fullmatch(r"[a-z][a-z0-9_]{0,49}", key):
        raise ValueError("Mã tác vụ chỉ gồm chữ thường, số và _, bắt đầu bằng chữ (ví dụ: van_ban_moi).")
    return key


def _validate_sheet_name(sheet_name: str) -> str:
    name = str(sheet_name or "").strip()
    if not name or len(name) > 31 or re.search(r"[\\/:*?\[\]]", name):
        raise ValueError("Tên sheet Excel không được rỗng, tối đa 31 ký tự và không chứa \\ / : * ? [ ].")
    return name


def _task_value_span(source: str, task_key: str, field: str):
    _key, call = _task_call(source, task_key)
    for keyword in call.keywords:
        if keyword.arg == field:
            return _offset(source, keyword.value.lineno, keyword.value.col_offset), _offset(
                source, keyword.value.end_lineno, keyword.value.end_col_offset
            )
    return None, call


def update_task_fields(task_key: str, updates: Dict[str, Any]) -> None:
    """Cập nhật field TaskConfig qua AST, nên selector có ngoặc/JSON vẫn an toàn."""
    task_key = _validate_task_key(task_key)
    source = CONFIG_PATH.read_text(encoding="utf-8")
    # Kiem tra task ton tai truoc khi ghi bat ky thay doi nao.
    _task_call(source, task_key)

    for field, value in updates.items():
        render = TASK_FIELD_RENDER.get(field)
        if render is None:
            continue
        if field == "navigation_steps" and not isinstance(value, list):
            raise ValueError("navigation_steps phải là một mảng JSON.")
        if field == "sheet_name":
            value = _validate_sheet_name(value)
            tasks = _tasks_dict(source)
            for other_key_node, other_call in zip(tasks.keys, tasks.values):
                if not isinstance(other_key_node, ast.Constant) or other_key_node.value == task_key:
                    continue
                if not isinstance(other_call, ast.Call):
                    continue
                for keyword in other_call.keywords:
                    if keyword.arg == "sheet_name" and isinstance(keyword.value, ast.Constant):
                        if str(keyword.value.value) == value:
                            raise ValueError(f"Tên sheet '{value}' đang được một tác vụ khác dùng.")

        result = _task_value_span(source, task_key, field)
        if isinstance(result[0], int):
            start, end = result
            source = source[:start] + render(value) + source[end:]
            continue

        _key, call = _task_call(source, task_key)
        close_paren = _offset(source, call.end_lineno, call.end_col_offset) - 1
        before = source[:close_paren].rstrip()
        separator = "\n" if before.endswith(",") else ",\n"
        source = before + separator + f"        {field}={render(value)},\n    " + source[close_paren:]

    # ast.parse truoc khi ghi de tranh tao config.py hong khi co du lieu bat thuong.
    ast.parse(source)
    CONFIG_PATH.write_text(source, encoding="utf-8")


def add_task(task_data: Dict[str, Any]) -> str:
    """Them TaskConfig moi vao TASKS. Khong dong den workbook/sheet cu."""
    key = _validate_task_key(task_data.get("key", ""))
    label = str(task_data.get("label", "")).strip()
    if not label:
        raise ValueError("Cần nhập tên hiển thị của tác vụ.")
    navigation_steps = task_data.get("navigation_steps", [])
    use_advanced_navigation = bool(task_data.get("use_advanced_navigation", False))
    if not isinstance(navigation_steps, list):
        raise ValueError("navigation_steps phải là một mảng JSON.")
    if use_advanced_navigation and not navigation_steps:
        raise ValueError("Đã bật điều hướng nâng cao nên cần ít nhất một bước JSON.")
    sidebar_item = str(task_data.get("sidebar_item", "")).strip()
    list_link = str(task_data.get("list_link", "")).strip()
    direct_href = str(task_data.get("list_link_href", "")).strip()
    tab_name = str(task_data.get("tab_name", "")).strip()
    extract_mode = str(task_data.get("extract_mode", "directive")).strip() or "directive"
    if not use_advanced_navigation and not direct_href and (not sidebar_item or not list_link):
        raise ValueError("Tác vụ mới cần chọn Đường dẫn trực tiếp hoặc nhập cả Sidebar và Tiểu mục.")
    if not use_advanced_navigation and extract_mode == "directive" and not tab_name:
        raise ValueError("Tác vụ dạng xử lý cần nhập Tab phải chọn.")

    sheet_name = _validate_sheet_name(task_data.get("sheet_name") or label)
    source = CONFIG_PATH.read_text(encoding="utf-8")
    tasks = _tasks_dict(source)
    existing_keys = [node.value for node in tasks.keys if isinstance(node, ast.Constant) and isinstance(node.value, str)]
    if key in existing_keys:
        raise ValueError(f"Mã tác vụ '{key}' đã tồn tại.")

    existing_sheets = set()
    for existing_key in existing_keys:
        _key, call = _task_call(source, existing_key)
        for keyword in call.keywords:
            if keyword.arg == "sheet_name" and isinstance(keyword.value, ast.Constant):
                existing_sheets.add(str(keyword.value.value))
    if sheet_name in existing_sheets:
        raise ValueError(f"Tên sheet '{sheet_name}' đang được một tác vụ khác dùng.")

    max_order = 0
    for existing_key in existing_keys:
        _key, call = _task_call(source, existing_key)
        for keyword in call.keywords:
            if keyword.arg == "sheet_order" and isinstance(keyword.value, ast.Constant):
                max_order = max(max_order, int(keyword.value.value))

    values = {
        "key": key,
        "label": label,
        "enabled": bool(task_data.get("enabled", True)),
        "role_pattern": str(task_data.get("role_pattern", "")).strip(),
        "sidebar_item": sidebar_item,
        "list_link": list_link,
        "list_link_href": direct_href,
        "tab_name": tab_name or None,
        "navigation_steps": navigation_steps,
        "use_advanced_navigation": use_advanced_navigation,
        "document_row_selector": str(task_data.get("document_row_selector", "tr.mat-row")).strip() or "tr.mat-row",
        "document_click_selector": str(task_data.get("document_click_selector", "")).strip(),
        "extract_mode": extract_mode,
        "max_documents": max(1, int(task_data.get("max_documents", 20))),
        "enable_download_pdf": bool(task_data.get("enable_download_pdf", True)),
        "enable_finish": bool(task_data.get("enable_finish", False)),
        "download_subdir": str(task_data.get("download_subdir") or f"VB_{key}").strip(),
        "sheet_name": sheet_name,
        "title_text": str(task_data.get("title_text") or f"TỔNG HỢP {label.upper()}").strip(),
        "sheet_order": max_order + 1,
        "debug_prefix": str(task_data.get("debug_prefix") or key).strip(),
    }
    rendered = [
        f"        {name}={TASK_FIELD_RENDER[name](value)},"
        for name, value in values.items()
        if name in TASK_FIELD_RENDER
    ]
    # key/debug_prefix khong nam trong TASK_FIELD_RENDER vi khong sua qua form cap nhat.
    rendered.insert(0, f"        key={json.dumps(key, ensure_ascii=False)},")
    rendered.append(f"        debug_prefix={json.dumps(values['debug_prefix'], ensure_ascii=False)},")
    block = f'    "{key}": TaskConfig(\n' + "\n".join(rendered) + "\n    ),\n"

    insert_at = _offset(source, tasks.end_lineno, tasks.end_col_offset) - 1
    source = source[:insert_at] + block + source[insert_at:]
    ast.parse(source)
    CONFIG_PATH.write_text(source, encoding="utf-8")
    return key


def remove_task(task_key: str) -> None:
    """Bo mot task khoi config.py; tuyet doi khong xoa sheet Excel hay du lieu."""
    task_key = _validate_task_key(task_key)
    source = CONFIG_PATH.read_text(encoding="utf-8")
    key_node, call = _task_call(source, task_key)
    start = _offset(source, key_node.lineno, key_node.col_offset)
    end = _offset(source, call.end_lineno, call.end_col_offset)
    while end < len(source) and source[end] in " \t":
        end += 1
    if end < len(source) and source[end] == ",":
        end += 1
    if end < len(source) and source[end] == "\n":
        end += 1
    source = source[:start] + source[end:]
    ast.parse(source)
    CONFIG_PATH.write_text(source, encoding="utf-8")


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
        # Gia tri nguoi dung tu dien qua trang Cai dat (rong = dang dung mac dinh/env).
        DISPLAY_BASE_DIR_OVERRIDE=_get(base_cfg, "DISPLAY_BASE_DIR_OVERRIDE"),
        DISPLAY_BASE_URL_OVERRIDE=_get(base_cfg, "DISPLAY_BASE_URL_OVERRIDE"),
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
