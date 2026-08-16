"""
Dinh nghia cau truc mot "tac vu" DOffice (VD: Van ban Chu tri, Van ban Phoi hop,
Cong viec Dang doan Phoi hop...).

Moi tac vu la 1 TaskConfig. File config.py se khai bao cac TaskConfig cu the trong
dict TASKS - do la noi duy nhat can chinh khi:
- doi vai tro (role) tren DOffice cua tung tac vu,
- doi so luong van ban toi da,
- bat/tat tac vu khoi lan chay "--all",
- doi ten sheet/tieu de Excel...

Khong can dong vao code logic trong cac module khac.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class TaskConfig:
    # --- Dinh danh ---
    key: str  # khoa noi bo, dung trong --tasks va trong dict TASKS, vd "chu_tri"
    label: str  # ten hien thi tieng Viet, vd "Van ban Chu tri - Da xu ly"
    enabled: bool = True  # co nam trong lua chon "Tat ca" / menu mac dinh khong

    # --- Chon dung chuc danh (role) tren DOffice ---
    # De trong "" neu tai khoan chi co 1 vai tro, khong can doi.
    # Co the dung nguyen van ten hien thi trong menu, vi du:
    #   "Pho Truyen tai dien", "Truong Truyen tai dien", "Chi bo 1", "Ky thuat"...
    # Gia tri nay duoc dung truc tiep nhu 1 regex (khong phan biet hoa/thuong),
    # nen ky tu dac biet cua regex (. * ? ( ) ...) se co nghia dac biet neu co.
    role_pattern: str = ""

    # --- Dieu huong menu DOffice ---
    sidebar_item: str = "Văn bản"  # "Văn bản" hoặc "Công việc"
    list_link: str = "Chờ xử lý"  # "Chờ xử lý" | "Đã xử lý" | "Chờ thực hiện"
    # Tab bat buoc voi extract_mode="directive". Ten co the la Chu tri/Phoi hop
    # hoac Chua xu ly/Da xu ly tuy man hinh DOffice. published co the khong co tab.
    tab_name: Optional[str] = None
    # Neu dat, sau khi chon role se goto thang href nay va bo qua Sidebar/Tieu
    # muc. Day la cach uu tien, dong thoi giai quyet text bi TRUNG (vd hai link
    # "Xem de biet"). Chap nhan path /... hoac URL day du cung origin DOffice.
    list_link_href: str = ""

    # --- Chuoi dieu huong co the tuy chinh ---
    # Neu co du lieu, navigation_steps se thay the luong sidebar_item/list_link/
    # tab_name co dinh ben tren. Moi buoc la dict, vi du:
    # {
    #   "name": "Mo Cong viec", "type": "sidebar", "value": "Cong viec"
    # }
    # {
    #   "name": "Chon bo loc", "type": "selector",
    #   "selector": "button[data-testid='filter']", "action": "click"
    # }
    # type ho tro: sidebar, link_text, link_href, tab, selector.
    # action cua selector: click (mac dinh) hoac wait. Co the them optional,
    # timeout_ms, delay_ms, nth va wait_for. Xem trang Cai dat de co mau JSON.
    # De [] chi dung cho cau hinh cu; he thong se tu dung 3 truong legacy o tren.
    navigation_steps: List[Dict[str, Any]] = field(default_factory=list)

    # None giu tuong thich voi task da tao o ban truoc: co navigation_steps thi
    # hieu la dang dung nang cao. False = dung 3 truong Sidebar/Tieu muc/Tab don
    # gian; True = chay navigation_steps JSON.
    use_advanced_navigation: Optional[bool] = None

    # Selector danh sach va phan tu can click ben trong mot dong. Hai truong nay
    # cho phep phu hop voi tai khoan/giao dien DOffice khac ma khong sua Python.
    document_row_selector: str = "tr.mat-row"
    document_click_selector: str = ""

    # --- Cach trich xuat du lieu tu dong danh sach ---
    #   "directive" (mac dinh): van ban chi dao (Chu tri/Phoi hop/Dang doan) - co
    #       khoi nguoi/noi dung chi dao, chu tri, phoi hop (cau truc div.vb-item).
    #   "published": van ban da ky duyet phat hanh (VB di) - cau truc DOM khac han,
    #       co Nguoi/Don vi soan thao thay cho khoi chi dao. Xem do_auto/extract.py.
    extract_mode: str = "directive"

    # --- Hanh vi xu ly van ban ---
    # Tac vu nay co bao gio can bam "Ket thuc nhanh" khong (vd chu_tri/vb_duyet:
    # van ban DA xu ly/DA phat hanh roi, khong bao gio bam Ket thuc). Dat False de
    # trang web Cai dat AN/LAM MO 2 o "Bam Ket thuc..." va "Hoi xac nhan..." cho
    # tac vu nay - tranh nguoi dung bat nham 1 hanh vi khong co tac dung/khong
    # dung. Khong lien quan enable_finish (gia tri THUC TE dang bat/tat).
    supports_finish: bool = True

    # Co bam "Ket thuc nhanh" + "Luu" sau khi tai PDF khong.
    # Danh sach "Da xu ly" (chu_tri) khong can vi van ban da xong roi.
    enable_finish: bool = True
    ask_confirm_before_finish: bool = False
    # Sau khi Ket thuc, DOffice thuong tu day van ban ke tiep len dong dau -> chon lai dong 0.
    always_pick_first_row_after_finish: bool = True
    # True: uu tien click icon co (chi co o danh sach "Chu tri - Da xu ly").
    prefer_flag_icon: bool = False
    # True (chu_tri): trich thong tin + kiem tra trung TRUOC khi mo van ban (tiet kiem thoi
    #   gian vi van ban da xu ly roi, khong can mo neu da co trong Excel).
    # False (phoi_hop, dang_doan): phai mo van ban ra thi thong tin chi dao moi hien du,
    #   nen mo truoc roi moi trich + kiem tra trung.
    check_duplicate_before_open: bool = False

    # --- Gioi han & tai PDF ---
    max_documents: int = 30
    enable_download_pdf: bool = True

    # --- Noi luu ---
    download_subdir: str = "VB_Khac"  # thu muc con duoi DOWNLOAD_BASE_DIR
    sheet_name: str = "Sheet"  # ten sheet trong file Excel gop
    title_text: str = "TỔNG HỢP VĂN BẢN"  # tieu de lon o dong 1 cua sheet
    sheet_order: int = 99  # thu tu sheet trong workbook (1, 2, 3...)
    debug_prefix: str = "doffice"  # tien to ten file anh debug khi loi


@dataclass
class TaskResult:
    key: str
    label: str
    ok: bool
    processed: int
    note: str = ""
    # Tom tat cac van ban MOI da ghi thanh cong trong lan chay nay (so_vb,
    # ngay_vb, trich_yeu) - dung de gui tin nhan tong ket qua Telegram
    # (xem do_auto/notify.py). Rong neu processed=0 hoac ok=False.
    documents: List[Dict[str, str]] = field(default_factory=list)
