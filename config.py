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

# ----------------------------------------------------------------------------------
# 1) CAU HINH CHUNG - ap dung cho moi tac vu
# ----------------------------------------------------------------------------------

# Trang goc DOffice. Cac tac vu se tu dieu huong bang menu tu day.
DOFFICE_URL = "https://doffice.npt.com.vn/"

# Phien dang nhap Playwright da luu sau khi chay login_save_state.py.
AUTH_STATE = Path("playwright/.auth/state.json")

# Thu muc goc chua PDF tai ve. Moi tac vu se tu tao 1 thu muc con ben trong
# (xem download_subdir trong tung TaskConfig ben duoi), vi du:
#   DOWNLOAD_BASE_DIR / "VB_Chu_tri_da_XL"
#   DOWNLOAD_BASE_DIR / "VB_Phoi_hop"
#   DOWNLOAD_BASE_DIR / "VB_Dang_doan_phoi_hop"
# Uu tien bien moi truong DOFFICE_DATA_DIR (dat khi chay trong Docker tren NAS,
# tro toi /data - la thu muc /volume1/homes/binhnx/Working/Van_ban duoc mount vao
# container). Neu khong co bien do (vd dang chay truc tiep tren Windows) thi
# quay ve duong dan Windows cu, nen file nay dung duoc o ca 2 moi truong.
DOWNLOAD_BASE_DIR = Path(os.environ.get("DOFFICE_DATA_DIR") or r"D:\OneDrive - NPT\9. Jobs\Van_ban")

# Duong dan HIEN THI trong Excel/web (cot "Thu muc luu" va hyperlink ten file) de
# nguoi dung bam MO DUOC file tren may Windows cua minh. Ly do: khi chay trong
# Docker tren NAS, file that nam o /data BEN TRONG container - duong dan do vo
# nghia voi may Windows. Tren cac may cua ban, thu muc du lieu tren NAS duoc map
# vao o S:, nen phai ghi vao Excel duong dan kieu "S:\..." thi click moi mo ra.
# Uu tien bien moi truong DOFFICE_DISPLAY_DIR (dat trong docker-compose.yml). Khi
# chay TRUC TIEP tren Windows (khong dat bien) thi mac dinh = DOWNLOAD_BASE_DIR
# nen khong thay doi gi so voi truoc.
DISPLAY_BASE_DIR = os.environ.get("DOFFICE_DISPLAY_DIR") or str(DOWNLOAD_BASE_DIR)

# 1 file Excel duy nhat, gom du lieu cua ca 3 tac vu, moi tac vu 1 sheet rieng
# (xem sheet_name trong tung TaskConfig ben duoi).
EXCEL_FILE = DOWNLOAD_BASE_DIR / "Tong_hop_DOffice.xlsx"

# File SQLite luu lich su cac lan chay (dung cho trang "Lich su" tren web dashboard).
# Tu tao khi chay lan dau, khong can tao truoc.
HISTORY_DB = DOWNLOAD_BASE_DIR / "doffice_auto_history.sqlite3"

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

TASKS = {
    "chu_tri": TaskConfig(
        key="chu_tri",
        label="Văn bản Chủ trì - Đã xử lý",
        enabled=True,
        role_pattern="Phó Truyền tải điện",
        sidebar_item="Văn bản",
        list_link="Đã xử lý",
        tab_name="Chủ trì",
        enable_finish=False,  # van ban da xu ly roi, khong can bam Ket thuc
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=False,
        prefer_flag_icon=True,  # danh sach nay co icon co, click on dinh hon
        check_duplicate_before_open=True,  # kiem tra trung TRUOC khi mo van ban, do tiet kiem thoi gian
        max_documents=20,
        enable_download_pdf=True,
        download_subdir="VB_Chu_tri_da_XL",
        sheet_name="Chủ trì",
        title_text="TỔNG HỢP CÁC VĂN BẢN CHỦ TRÌ ĐÃ XỬ LÝ",
        sheet_order=1,
        debug_prefix="chu_tri_da_xl",
    ),
    "phoi_hop": TaskConfig(
        key="phoi_hop",
        label="Văn bản Phối hợp - Chờ xử lý",
        enabled=True,
        role_pattern="Phó Truyền tải điện",
        sidebar_item="Văn bản",
        list_link="Chờ xử lý",
        tab_name="Phối hợp",
        enable_finish=True,
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=True,
        prefer_flag_icon=False,
        check_duplicate_before_open=False,  # phai mo van ban thi thong tin chi dao moi hien du
        max_documents=50,
        enable_download_pdf=True,
        download_subdir="VB_Phoi_hop",
        sheet_name="Phối hợp",
        title_text="TỔNG HỢP CÁC VĂN BẢN PHỐI HỢP",
        sheet_order=2,
        debug_prefix="vb_phoi_hop",
    ),
    "dang_doan": TaskConfig(
        key="dang_doan",
        label="Công việc Đảng/Đoàn/Công đoàn - Chờ thực hiện (Phối hợp)",
        enabled=True,
        role_pattern="Chi bộ 1",
        sidebar_item="Công việc",
        list_link="Chờ thực hiện",
        tab_name="Phối hợp",
        enable_finish=True,
        ask_confirm_before_finish=False,
        always_pick_first_row_after_finish=True,
        prefer_flag_icon=False,
        check_duplicate_before_open=False,
        max_documents=50,
        enable_download_pdf=True,
        download_subdir="VB_Dang_doan_phoi_hop",
        sheet_name="Đảng - Đoàn",
        title_text="TỔNG HỢP CÁC VĂN BẢN ĐẢNG, CÔNG ĐOÀN ĐÃ XỬ LÝ",
        sheet_order=3,
        debug_prefix="dang_doan_phoi_hop",
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
ENABLE_TELEGRAM_NOTIFY = True
TELEGRAM_BOT_TOKEN = "8825591535:AAGMvHrdPszpPBwKEvPI1hRLZtVewfRF6KM"
TELEGRAM_CHAT_ID = "-5528425169"
TELEGRAM_NOTIFY_ONLY_IF_NEW = False
