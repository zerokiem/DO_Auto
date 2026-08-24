# ==================================================================================
# CAU HINH DOffice Auto - FILE NAY LA NOI DUY NHAT CAN CHINH KHI DUNG TREN MAY MOI
# HOAC CHO NGUOI DUNG KHAC.
#
# Neu ban la nguoi dung moi (khong phai Binh):
#   1) Doi ROLE_PATTERN cua tung tac vu ben duoi (TASKS) thanh dung ten chuc danh
#      hien thi tren DOffice cua ban (xem huong dan trong README.md muc 4).
#   2) Doi DOWNLOAD_BASE_DIR / AUTH_STATE neu muon luu o vi tri khac.
#   3) Chay python login_save_state.py 1 lan de tao phien dang nhap rieng.
# ==================================================================================
import os
from pathlib import Path

from do_auto.task_types import TaskConfig

APP_VERSION = "1.1.3"
BROWSER_VIEWPORT = {"width": 1280, "height": 720}
PROJECT_DIR = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------------
# 1) CAU HINH CHUNG - ap dung cho moi tac vu
# ----------------------------------------------------------------------------------

# Trang goc DOffice. Cac tac vu se tu dieu huong bang menu tu day.
DOFFICE_URL = "https://doffice.npt.com.vn/"

# Phien dang nhap Playwright da luu sau khi chay login_save_state.py.
# Neo vao thu muc project de web dashboard, CLI va Task Scheduler luon dung
# chung mot file, bat ke thu muc hien tai cua tien trinh.
_AUTH_STATE_VALUE = os.environ.get("DOFFICE_AUTH_STATE", "playwright/.auth/state.json")
_AUTH_STATE_PATH = Path(_AUTH_STATE_VALUE).expanduser()
AUTH_STATE = _AUTH_STATE_PATH if _AUTH_STATE_PATH.is_absolute() else PROJECT_DIR / _AUTH_STATE_PATH

# Thu muc goc chua PDF tai ve. Moi tac vu se tu tao 1 thu muc con ben trong
# (xem download_subdir trong tung TaskConfig ben duoi), vi du:
#   DOWNLOAD_BASE_DIR / "VB_Chu_tri_da_XL"
#   DOWNLOAD_BASE_DIR / "VB_Phoi_hop"
#   DOWNLOAD_BASE_DIR / "VB_Dang_doan_phoi_hop"
# Thu muc luu PDF, Excel, lich su va log. Co the dat ngay tren trang
# Cai dat. Thu tu uu tien: gia tri dat tren web > DOFFICE_DATA_DIR > ./data.
#
# Luu y Docker/NAS: day la duong dan NHIN THAY BOI container (thuong /data),
# khong phai duong dan host/NAS. Muon doi thu muc tren NAS thi doi
# DOFFICE_DATA_HOST trong .env roi chay lai docker compose.
DOWNLOAD_BASE_DIR_OVERRIDE = ""
DOWNLOAD_BASE_DIR = Path(
    DOWNLOAD_BASE_DIR_OVERRIDE
    or os.environ.get("DOFFICE_DATA_DIR")
    or (Path(__file__).resolve().parent / "data")
)

# Link mo PDF duoc tu dong tao tu dung thu muc luu o Windows (file://). Khi chay
# Docker/NAS, dat URL public cua dashboard + /vb trong .env de link mo qua web,
# vi du http://100.100.1.254:8877/vb hoac URL Tailscale tuong ung.
DISPLAY_BASE_URL = os.environ.get("DOFFICE_DISPLAY_URL") or ""

# 1 file Excel duy nhat, gom du lieu cua cac tac vu, moi tac vu 1 sheet rieng
# (xem sheet_name trong tung TaskConfig ben duoi).
EXCEL_FILE = DOWNLOAD_BASE_DIR / "Tong_hop_DOffice.xlsx"

# File SQLite luu lich su cac lan chay (dung cho trang "Lich su" tren web dashboard).
# Tu tao khi chay lan dau, khong can tao truoc.
HISTORY_DB = DOWNLOAD_BASE_DIR / "doffice_auto_history.jsonl"

# Thu muc luu file log DAY DU cua TUNG LAN CHAY (moi lan chay run_doffice.py hoac
# tu web deu tu tao 1 file .log o day, khong chi rieng luc chay bang Scheduled Task).
LOGS_DIR = DOWNLOAD_BASE_DIR / "logs"

# Khoa kiem tra trung, dung chung cho ca 3 tac vu:
# - "so_vb_ngay_vb": So VB + Ngay VB. Khuyen nghi dung mac dinh.
# - "so_vb": chi dung So VB. Chi dung khi chac chan So VB khong lap.
# - "so_vb_time": So VB + Thoi gian chi dao. Dung khi can phan biet nhieu chi dao
#   tren cung 1 van ban.
DUPLICATE_CHECK_MODE = "so_vb_ngay_vb"

# Khi gap van ban da co trong Excel thi dung tac vu do lai (khuyen nghi True vi
# danh sach DOffice thuong sap xep moi nhat len tren, gap dong trung nghia la cac
# dong duoi cung da tong hop roi).
STOP_WHEN_DUPLICATE_FOUND = True

# Sau moi REFRESH_LIST_EVERY van ban da GHI THANH CONG thi load lai danh sach tu
# dau. 0 = khong tu reload (khuyen nghi de mac dinh 0 tru khi thay danh sach ao
# cua DOffice bi loi/mat dong sau khi xu ly nhieu van ban).
REFRESH_LIST_EVERY = 0

# Timeout bat su kien download theo tung lan thu (mili giay). Lan 1 co y ngan vi
# DOffice/PDF viewer doi khi click dau chi kich hoat viewer, chua phat download
# ngay; cac lan sau tang dan de du thoi gian cho file lon.
DOWNLOAD_ATTEMPT_TIMEOUTS_MS = [7000, 15000, 20000]

# Toc do thao tac trinh duyet, de quan sat khi test. Dat 0 khi chay that/chay
# scheduled task de nhanh hon.
SLOW_MO_MS = 250

# Co dung cho nhan Enter truoc khi dong browser khong. Nen de False khi chay
# scheduled task (khong co ai bam Enter se lam treo task).
PAUSE_BEFORE_CLOSE = False

# Chi dien neu trang DOffice co NHIEU HON 1 nut header dang "<Ten> Phòng ban:"
# (rat hiem gap). Mac dinh de trong "" - luc do he thong se tim bat ky nut nao
# ket thuc bang "Phòng ban:", nen dung duoc voi bat ky tai khoan DOffice nao dang
# nhap ma khong can go ten.
ROLE_BUTTON_NAME_HINT = ""


# ----------------------------------------------------------------------------------
# 2) CAU HINH RIENG TUNG TAC VU
#
#    role_pattern: go dung ten chuc danh hien thi trong menu "... Phòng ban:" tren
#    DOffice cua ban. Vi du menu co "Phó Truyền tải điện Nguyễn Xuân Bình" thi go
#    "Phó Truyền tải điện" (khong can go het ten). Neu tai khoan chi co 1 vai tro
#    (khong hien menu chon) thi de "".
# ----------------------------------------------------------------------------------

LEGACY_TASKS = {
    "chu_tri": TaskConfig(
        key="chu_tri",
        label="Văn bản Chủ trì - Đã xử lý",
        enabled=True,
        role_pattern="Phó Truyền tải điện",
        sidebar_item="Văn bản",
        list_link="Đã xử lý",
        tab_name="Chủ trì",
        list_link_href="",
        supports_finish=False,  # khong bao gio dung Ket thuc - an mo 2 o do tren web Cai dat
        enable_finish=False,  # van ban da xu ly roi, khong can bam Ket thuc
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=False,
        prefer_flag_icon=True,  # danh sach nay co icon co, click on dinh hon
        check_duplicate_before_open=True,  # kiem tra trung TRUOC khi mo van ban, do tiet kiem thoi gian
        max_documents=147,
        enable_download_pdf=True,
        download_subdir="VB_Chu_tri_da_XL",
        sheet_name="Chủ trì",
        title_text="TỔNG HỢP CÁC VĂN BẢN CHỦ TRÌ ĐÃ XỬ LÝ",
        sheet_order=1,
        navigation_steps=[{'name': 'Mở Văn bản', 'type': 'sidebar', 'value': 'Văn bản'},
 {'name': 'Mở Đã xử lý', 'type': 'link_text', 'value': 'Đã xử lý'},
 {'empty_if_zero': True, 'name': 'Chọn tab Chủ trì', 'type': 'tab', 'value': 'Chủ trì'}],
        use_advanced_navigation=False,
        debug_prefix="chu_tri_da_xl",
        document_row_selector="tr.mat-row",
        document_click_selector="",
        extract_mode="directive",
    ),
    "phoi_hop": TaskConfig(
        key="phoi_hop",
        label="Văn bản Phối hợp - Chờ xử lý",
        enabled=True,
        role_pattern="Phó Truyền tải điện",
        sidebar_item="Văn bản",
        list_link="Chờ xử lý",
        tab_name="Phối hợp",
        list_link_href="/congviec/ld-xly-vb/ChoXL",
        enable_finish=True,
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=True,
        prefer_flag_icon=False,
        check_duplicate_before_open=False,  # phai mo van ban thi thong tin chi dao moi hien du
        max_documents=147,
        enable_download_pdf=True,
        download_subdir="VB_Phoi_hop",
        sheet_name="Phối hợp",
        title_text="TỔNG HỢP CÁC VĂN BẢN PHỐI HỢP",
        sheet_order=2,
        navigation_steps=[{'name': 'Mở Văn bản', 'type': 'sidebar', 'value': 'Văn bản'},
 {'name': 'Mở Chờ xử lý', 'type': 'link_text', 'value': 'Chờ xử lý'},
 {'empty_if_zero': True, 'name': 'Chọn tab Phối hợp', 'type': 'tab', 'value': 'Phối hợp'}],
        use_advanced_navigation=False,
        debug_prefix="vb_phoi_hop",
        document_row_selector="tr.mat-row",
        document_click_selector="",
        extract_mode="directive",
    ),
    "dang_doan": TaskConfig(
        key="dang_doan",
        label="Công việc Đảng / Công đoàn - Phối hợp",
        enabled=True,
        role_pattern="Chi bộ 1",
        sidebar_item="Công việc",
        list_link="Chờ thực hiện",
        tab_name="Phối hợp",
        list_link_href="/congviec/cviec-cvien-xuly/chothuchien",
        enable_finish=True,
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=True,
        prefer_flag_icon=False,
        check_duplicate_before_open=False,
        max_documents=147,
        enable_download_pdf=True,
        download_subdir="VB_Dang_doan_phoi_hop",
        sheet_name="Đảng - Đoàn",
        title_text="TỔNG HỢP CÁC VĂN BẢN ĐẢNG, CÔNG ĐOÀN ĐÃ XỬ LÝ",
        sheet_order=3,
        navigation_steps=[{'name': 'Mở Công việc', 'type': 'sidebar', 'value': 'Công việc'},
 {'name': 'Mở Chờ thực hiện', 'type': 'link_text', 'value': 'Chờ thực hiện'},
 {'empty_if_zero': True, 'name': 'Chọn tab Phối hợp', 'type': 'tab', 'value': 'Phối hợp'}],
        use_advanced_navigation=False,
        debug_prefix="dang_doan_phoi_hop",
        document_row_selector="tr.mat-row",
        document_click_selector="",
        extract_mode="directive",
    ),
    "vb_duyet": TaskConfig(
        key="vb_duyet",
        label="Văn bản đã duyệt - Phát hành",
        enabled=True,
        role_pattern="Phó Truyền tải điện",
        sidebar_item="Văn bản",
        list_link="Đã phát hành",
        # "Đã phát hành" co ca vbdi (van ban di) lan vbnb (noi bo) -> phai dieu
        # huong bang href chinh xac, khong dung text link (se bi trung). Day la VB
        # di (đã ký duyệt phát hành ra ngoai).
        list_link_href="/duthaovanban/danhsach/vbdi/phathanh",
        tab_name=None,
        extract_mode="published",  # cau truc DOM khac VB chi dao, xem extract.py
        supports_finish=False,  # khong bao gio dung Ket thuc - an mo 2 o do tren web Cai dat
        enable_finish=False,  # van ban da phat hanh roi, khong bam Ket thuc
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=False,
        prefer_flag_icon=False,
        check_duplicate_before_open=True,  # thong tin hien du tren dong danh sach
        max_documents=149,  # tang len neu can tai bu backlog lan dau (vd qua trang Cai dat)
        enable_download_pdf=True,
        download_subdir="VB_duyet",
        sheet_name="Văn bản đã duyệt",
        title_text="TỔNG HỢP CÁC VĂN BẢN ĐÃ KÝ DUYỆT PHÁT HÀNH",
        sheet_order=4,
        navigation_steps=[{'name': 'Mở Văn bản', 'type': 'sidebar', 'value': 'Văn bản'},
 {'href': '/duthaovanban/danhsach/vbdi/phathanh', 'name': 'Mở Đã phát hành', 'type': 'link_href'}],
        use_advanced_navigation=False,
        debug_prefix="vb_duyet",
        document_row_selector="tr.mat-row",
        document_click_selector="",
    ),
        "vb_de_biet": TaskConfig(
        key="vb_de_biet",
        label="Văn bản - Xem để biết",
        enabled=True,
        role_pattern="",
        sidebar_item="Văn bản",
        list_link="Xem để biết",
        list_link_href="",
        tab_name="Chưa xử lý",
        navigation_steps=[],
        use_advanced_navigation=False,
        document_row_selector="tr.mat-row",
        document_click_selector="",
        # Xem de biet dung cau truc dong Cong viec (KY_HIEU_CV), khong phai
        # cau truc Van ban da phat hanh. Dung directive de click So VB va tach
        # rieng So VB/Ngay/Noi phat hanh/Trich yeu.
        extract_mode="directive",
        max_documents=2,
        enable_download_pdf=True,
        enable_finish=True,
        download_subdir="VB_de_biet",
        sheet_name="de_biet",
        title_text="TỔNG HỢP VĂN BẢN - XEM ĐỂ BIẾT",
        sheet_order=5,
        debug_prefix="vb_de_biet",
        ask_confirm_before_finish=False,
    ),
}


# ----------------------------------------------------------------------------------
# TAC VU MAC DINH CHO BAN PHAT HANH MOI. Cac tac vu cu o tren duoc giu trong file
# de nguoi dang nang cap co the tham chieu, nhung khong nam trong danh sach chay.
TASKS = {
    "vb_phoi_hop": TaskConfig(
        key="vb_phoi_hop",
        label="Văn bản - Phối hợp",
        role_pattern="Tổ Tổng hợp",
        list_link_href="/congviec/cviec-cvien-xuly/chothuchien",
        tab_name="Phối hợp",
        extract_mode="directive",
        max_documents=2,
        enable_download_pdf=True,
        enable_finish=True,
        download_subdir="VB_phoi_hop",
        sheet_name="Phối hợp",
        title_text="TỔNG HỢP VĂN BẢN - PHỐI HỢP",
        sheet_order=1,
        debug_prefix="vb_phoi_hop",
    ),
    "vb_xem_db": TaskConfig(
        key="vb_xem_db",
        label="Văn bản - Xem để biết",
        role_pattern="Tổ Tổng hợp",
        list_link_href="/congviec/cviec-cvien-xuly/xemdb",
        tab_name="Chưa xử lý",
        extract_mode="directive",
        max_documents=2,
        enable_download_pdf=True,
        enable_finish=True,
        download_subdir="VB_de_biet",
        sheet_name="Xem để biết",
        title_text="TỔNG HỢP VĂN BẢN - XEM ĐỂ BIẾT",
        sheet_order=2,
        debug_prefix="vb_xem_db",
    ),
}

# ----------------------------------------------------------------------------------
# 3) CHE DO TEST AN TOAN
#    Chay: python run_doffice.py --all --test
#    se tu dong ghi de cac gia tri ben duoi (khong sua o day, khong anh huong che
#    do chay that).
# ----------------------------------------------------------------------------------
TEST_MODE_OVERRIDES = {
    "max_documents": 1,
    "enable_finish": False,
    "ask_confirm_before_finish": True,
}
TEST_MODE_PAUSE_BEFORE_CLOSE = True
TEST_MODE_SLOW_MO_MS = 800

# --- Tự động thêm bởi trang web Cài đặt (biến cấu hình mới từ bản cập nhật) ---
ENABLE_TELEGRAM_NOTIFY = os.environ.get("DOFFICE_TELEGRAM_ENABLED", "").strip().lower() in {"1", "true", "yes", "on"}
TELEGRAM_BOT_TOKEN = os.environ.get("DOFFICE_TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("DOFFICE_TELEGRAM_CHAT_ID", "")
TELEGRAM_NOTIFY_ONLY_IF_NEW = os.environ.get("DOFFICE_TELEGRAM_ONLY_IF_NEW", "").strip().lower() in {"1", "true", "yes", "on"}
