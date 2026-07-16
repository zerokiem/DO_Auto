"""
Ghi log Excel gop - 1 file .xlsx, moi tac vu 1 sheet rieng
(vd: Chu_tri, Phoi_hop, Dang_doan).

Luu y quan trong so voi 3 script goc: 3 script goc moi script dung 1 file .xlsx
rieng nen ham init luon dung wb.active la an toan. Khi gop chung 1 file nhieu
sheet, KHONG duoc dung wb.active de tim/doi ten sheet nua (se lam hong sheet
dang active neu sheet muc tieu chua ton tai) - ham init_excel_log() ben duoi da
sua loi nay: neu sheet can dung chua co, se tao sheet moi (create_sheet) thay vi
doi ten sheet dang active.
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, Set

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import text_utils

HEADERS = [
    "STT",
    "Số VB",
    "Ngày VB",
    "Nơi phát hành",
    "Trích yếu",
    "Người chỉ đạo",
    "Thời gian chỉ đạo",
    "Nội dung chỉ đạo",
    "Chủ trì",
    "Phối hợp",
    "Thư mục lưu",
    "Tên file lưu",
    "Thời gian lưu",
]

TITLE_ROW = 1
HEADER_ROW = 2
DATA_START_ROW = 3

# Cot theo chi so 0-based trong openpyxl values_only iterator (0 = STT).
COL_SO_VB = 1
COL_NGAY_VB = 2
COL_THOI_GIAN_CHI_DAO = 6
COL_TEN_FILE_LUU = 11


def apply_excel_layout(ws, title_text: str) -> None:
    """Chuan hoa layout Excel: title, header, do rong cot, freeze, filter, an cot."""
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    title_fill = PatternFill("solid", fgColor="D9EAF7")
    title_font = Font(color="1F4E78", bold=True, size=14)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    first_row_values = [ws.cell(row=1, column=i).value for i in range(1, len(HEADERS) + 1)]
    if first_row_values == HEADERS:
        ws.insert_rows(1)

    ws.cell(row=TITLE_ROW, column=1).value = title_text
    try:
        for merged in list(ws.merged_cells.ranges):
            if str(merged).startswith("A1:"):
                ws.unmerge_cells(str(merged))
    except Exception:
        pass

    title_cell = ws.cell(row=TITLE_ROW, column=1)
    title_cell.fill = title_fill
    title_cell.font = title_font
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    title_cell.border = border
    ws.row_dimensions[TITLE_ROW].height = 28

    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col_idx)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[HEADER_ROW].height = 35

    widths = {1: 6, 2: 18, 3: 13, 4: 30, 5: 50, 6: 18, 7: 20, 8: 30, 9: 16, 10: 30, 11: 35, 12: 45, 13: 20}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # An cot G = Thoi gian chi dao. Can thi Unhide lai trong Excel.
    ws.column_dimensions["G"].hidden = True

    ws.freeze_panes = "D3"
    ws.auto_filter.ref = "A2:M2"

    for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _get_or_create_sheet(wb, sheet_name: str):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    # Neu workbook chi co dung 1 sheet mac dinh trong ("Sheet") thi doi ten sheet do
    # thay vi tao them sheet rong thua.
    if wb.sheetnames == ["Sheet"] and wb["Sheet"].max_row <= 1:
        ws = wb["Sheet"]
        ws.title = sheet_name
        return ws
    return wb.create_sheet(title=sheet_name)


def init_excel_log(excel_file: Path, sheet_name: str, title_text: str) -> None:
    """Dam bao file Excel va sheet muc tieu ton tai, dung layout. An toan khi goi
    nhieu lan / cho nhieu sheet khac nhau trong cung 1 workbook."""
    excel_file.parent.mkdir(parents=True, exist_ok=True)

    if excel_file.exists():
        wb = load_workbook(excel_file)
        ws = _get_or_create_sheet(wb, sheet_name)
        apply_excel_layout(ws, title_text)
        wb.save(excel_file)
        wb.close()
        return

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    apply_excel_layout(ws, title_text)
    wb.save(excel_file)
    wb.close()
    print(f"✅ Đã tạo file Excel tổng hợp: {excel_file}")


def ensure_all_sheets(excel_file: Path, tasks: Dict[str, "object"]) -> None:
    """Tao san tat ca sheet cho moi tac vu trong TASKS theo dung thu tu sheet_order,
    de thu tu tab trong Excel luon nhat quan du nguoi dung chay tac vu nao truoc."""
    excel_file.parent.mkdir(parents=True, exist_ok=True)
    ordered = sorted(tasks.values(), key=lambda t: t.sheet_order)

    if excel_file.exists():
        wb = load_workbook(excel_file)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    for task in ordered:
        ws = _get_or_create_sheet(wb, task.sheet_name)
        apply_excel_layout(ws, task.title_text)

    wb.save(excel_file)
    wb.close()


def get_next_stt(ws) -> int:
    if ws.max_row < DATA_START_ROW:
        return 1
    return max(1, ws.max_row - HEADER_ROW + 1)


def get_next_excel_stt(excel_file: Path, sheet_name: str) -> int:
    init_excel_log(excel_file, sheet_name, sheet_name)
    wb = load_workbook(excel_file, read_only=False)
    ws = wb[sheet_name]
    stt = get_next_stt(ws)
    wb.close()
    return stt


def file_uri_from_path(path_text: str) -> str:
    try:
        return Path(path_text).resolve().as_uri()
    except Exception:
        return "file:///" + path_text.replace("\\", "/").replace(" ", "%20")


def append_excel_log(excel_file: Path, sheet_name: str, title_text: str, data: Dict[str, str]) -> None:
    init_excel_log(excel_file, sheet_name, title_text)

    wb = load_workbook(excel_file)
    ws = wb[sheet_name]
    stt = get_next_stt(ws)

    row = [
        stt,
        data.get("so_vb", ""),
        data.get("ngay_vb", ""),
        data.get("noi_phat_hanh", ""),
        data.get("trich_yeu", ""),
        data.get("nguoi_chi_dao", ""),
        data.get("thoi_gian_chi_dao", ""),
        data.get("noi_dung_chi_dao", ""),
        data.get("chu_tri", ""),
        data.get("phoi_hop", ""),
        data.get("thu_muc_luu", ""),
        data.get("ten_file_luu", ""),
        data.get("thoi_gian_luu", ""),
    ]
    ws.append(row)

    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for cell in ws[ws.max_row]:
        cell.border = border
        cell.alignment = Alignment(vertical="top", wrap_text=True)

    saved_folder = data.get("thu_muc_luu", "")
    saved_name = data.get("ten_file_luu", "")
    if saved_folder and saved_name:
        file_cell = ws.cell(row=ws.max_row, column=12)
        saved_path = str(Path(saved_folder) / saved_name)
        file_cell.hyperlink = file_uri_from_path(saved_path)
        file_cell.style = "Hyperlink"
        file_cell.comment = None

    apply_excel_layout(ws, title_text)
    wb.save(excel_file)
    wb.close()
    print(f"✅ [{sheet_name}] Đã ghi Excel dòng STT {stt}: {data.get('so_vb', '')}")


def build_duplicate_key(data: Dict[str, str], duplicate_check_mode: str) -> str:
    so_vb = text_utils.normalize_key(data.get("so_vb", ""))
    ngay_vb = text_utils.normalize_key(data.get("ngay_vb", ""))
    if duplicate_check_mode == "so_vb_time":
        tg = text_utils.normalize_key(data.get("thoi_gian_chi_dao", ""))
        return f"{so_vb}|{tg}"
    if duplicate_check_mode == "so_vb":
        return so_vb
    return f"{so_vb}|{ngay_vb}"


def load_existing_duplicate_keys(excel_file: Path, sheet_name: str, duplicate_check_mode: str) -> Set[str]:
    init_excel_log(excel_file, sheet_name, sheet_name)
    keys: Set[str] = set()

    wb = load_workbook(excel_file, read_only=True, data_only=True)
    ws = wb[sheet_name]

    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        if not row or len(row) < 7:
            continue
        so_vb = str(row[COL_SO_VB] or "")
        ngay_vb = str(row[COL_NGAY_VB] or "")
        thoi_gian = str(row[COL_THOI_GIAN_CHI_DAO] or "")
        if not so_vb.strip():
            continue
        if duplicate_check_mode == "so_vb_time":
            keys.add(f"{text_utils.normalize_key(so_vb)}|{text_utils.normalize_key(thoi_gian)}")
        elif duplicate_check_mode == "so_vb":
            keys.add(text_utils.normalize_key(so_vb))
        else:
            keys.add(f"{text_utils.normalize_key(so_vb)}|{text_utils.normalize_key(ngay_vb)}")

    wb.close()
    return keys


def load_existing_filenames(excel_file: Path, sheet_name: str) -> Set[str]:
    init_excel_log(excel_file, sheet_name, sheet_name)
    names: Set[str] = set()
    wb = load_workbook(excel_file, read_only=True, data_only=True)
    ws = wb[sheet_name]
    for row in ws.iter_rows(min_row=DATA_START_ROW, values_only=True):
        filename = str(row[COL_TEN_FILE_LUU] or "") if row and len(row) > COL_TEN_FILE_LUU else ""
        if filename.strip():
            names.add(text_utils.normalize_key(filename))
    wb.close()
    return names
