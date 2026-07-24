"""
Dieu huong DOffice: dang nhap, chon chuc danh (role), vao dung sidebar/link/tab,
lay dong (row) van ban trong danh sach.

Gop tu 3 script goc. Diem khac nhau giua cac tac vu (sidebar "Van ban" hay
"Cong viec", link "Cho xu ly"/"Da xu ly"/"Cho thuc hien", ten tab...) duoc truyen
vao qua TaskConfig thay vi hard-code, de mot bo ham duy nhat dung duoc cho ca 3
tac vu (va de nguoi khac them tac vu moi neu can).
"""
from __future__ import annotations

import re

from . import text_utils
from .task_types import TaskConfig


def click_login_if_needed(page) -> None:
    """Neu DOffice dang hien nut Dang nhap thi bam; da vao san thi bo qua."""
    try:
        btn = page.get_by_role("button", name=re.compile(r"Đăng nhập", re.I))
        if btn.count() > 0:
            btn.first.click(timeout=5000)
            print("✅ Đã click Đăng nhập.")
            text_utils.wait(page, 2000)
    except Exception:
        pass


def choose_role_if_needed(page, role_pattern: str, name_hint: str = "") -> None:
    """Chon dung chuc danh/vai tro tren DOffice truoc khi vao danh sach.

    role_pattern duoc dung truc tiep nhu regex (khong phan biet hoa/thuong) de
    nguoi dung co the go nguyen van chuc danh hien thi tren DOffice cua ho, khong
    can sua code. De trong "" neu tai khoan chi co 1 vai tro (khong co menu chon).

    name_hint (tuy chon): 1 phan ten hien thi tren nut goc header, vd "Nguyễn Xuân
    Bình", chi can dien neu trang co nhieu hon 1 nut ket thuc bang "Phong ban:".
    Mac dinh de trong, nut se duoc tim theo mau chung "... Phòng ban:" nen dung
    duoc voi bat ky tai khoan nao dang nhap, khong rieng gi 1 nguoi.
    """
    if not role_pattern:
        return

    print("\n--- Kiểm tra/chọn đúng chức danh ---")
    try:
        hint_part = re.escape(name_hint) + r".*" if name_hint else ""
        btn_pattern = re.compile(hint_part + r"Phòng ban", re.I)
        btn = page.get_by_role("button", name=btn_pattern)
        btn.wait_for(state="visible", timeout=8000)
        btn.click()
        text_utils.wait(page, 500)

        role = page.get_by_role("menuitem", name=re.compile(role_pattern, re.I))
        role.wait_for(state="visible", timeout=5000)
        role.click()
        print(f"✅ Đã chọn chức danh khớp mẫu: {role_pattern}")
        text_utils.wait(page, 1500)
    except Exception as e:
        print("⚠️ Không chọn lại chức danh được hoặc không cần chọn:", e)


def click_sidebar_item(page, sidebar_item: str, debug_prefix: str) -> bool:
    try:
        page.locator("fuse-vertical-navigation").get_by_text(sidebar_item, exact=True).click(timeout=8000)
        print(f"✅ Click sidebar {sidebar_item} trong fuse-vertical-navigation.")
        return True
    except Exception as e:
        print(f"⚠️ Không click được sidebar {sidebar_item} bằng fuse navigation:", e)

    for n in (1, 3, 0, 2):
        try:
            page.locator("div").filter(has_text=sidebar_item).nth(n).click(timeout=6000)
            print(f"✅ Click sidebar {sidebar_item} bằng locator div filter nth({n}).")
            return True
        except Exception:
            continue

    print(f"❌ Không click được sidebar {sidebar_item}.")
    text_utils.save_debug(page, debug_prefix, "sidebar_failed")
    return False


def click_tab(page, tab_name: str) -> None:
    try:
        page.get_by_role("tab", name=re.compile(re.escape(tab_name), re.I)).click(timeout=10000)
        print(f"✅ Click tab {tab_name}.")
    except Exception as e:
        print(f"⚠️ Không click được tab {tab_name} bằng role, thử bằng text:", e)
        try:
            page.get_by_text(re.compile(re.escape(tab_name) + r"\s*\(\d+\)", re.I)).click(timeout=10000)
            print(f"✅ Click tab {tab_name} bằng text.")
        except Exception as e2:
            print(f"⚠️ Không click được tab {tab_name}, kiểm tra xem danh sách có sẵn không:", e2)


def get_tab_count(page, tab_name: str):
    """Doc so trong ngoac canh ten tab (vd 'Phối hợp (0)') de biet danh sach dang
    THUC SU RONG (0 van ban cho xu ly) hay khong - tranh nham voi loi that su khi
    khong thay row nao. Tra ve None neu khong doc duoc so (vd DOffice doi giao dien),
    de cho nhanh gia dinh la loi that nhu truoc gio."""
    try:
        el = page.get_by_text(re.compile(re.escape(tab_name) + r"\s*\(\d+\)", re.I)).first
        text = el.inner_text(timeout=3000)
        m = re.search(r"\((\d+)\)", text)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    return None


def open_task_list(page, task: TaskConfig, cfg):
    """Vao dung man hinh danh sach cua 1 tac vu: goto trang goc -> dang nhap neu
    can -> chon chuc danh -> click sidebar -> click link -> click tab (neu co).

    Tra ve True (co danh sach, it nhat 1 van ban), False (loi that su - khong vao
    duoc trang/danh sach), hoac None (danh sach RONG THAT SU - tab hien "(0)",
    khong phai loi, khong con van ban nao can xu ly)."""
    where = f"{task.sidebar_item} / {task.list_link}"
    if task.tab_name:
        where += f" / {task.tab_name}"
    print(f"\n--- Vào {where} ---")

    page.goto(cfg.DOFFICE_URL, wait_until="domcontentloaded")
    text_utils.wait(page, 2500)

    click_login_if_needed(page)
    choose_role_if_needed(page, task.role_pattern, cfg.ROLE_BUTTON_NAME_HINT)

    if not click_sidebar_item(page, task.sidebar_item, task.debug_prefix):
        return False
    text_utils.wait(page, 1000)

    # Uu tien dieu huong bang href chinh xac neu task khai bao (tranh loi "strict
    # mode violation" khi text link bi trung, vd "Đã phát hành" co ca vbdi + vbnb).
    if getattr(task, "list_link_href", ""):
        link_locator = page.locator(f'a[href="{task.list_link_href}"]').first
        link_desc = f"Link (href) {task.list_link_href}"
    else:
        link_locator = page.get_by_role("link", name=re.compile(re.escape(task.list_link), re.I))
        link_desc = f"Link {task.list_link}"

    if not text_utils.safe_click(link_locator, link_desc, timeout=10000):
        text_utils.save_debug(page, task.debug_prefix, "link_failed")
        return False
    text_utils.wait(page, 2000)

    if task.tab_name:
        click_tab(page, task.tab_name)
        text_utils.wait(page, 500)
        count = get_tab_count(page, task.tab_name)
        if count == 0:
            print(f"ℹ️ Tab {task.tab_name} hiện (0) văn bản - không còn gì cần xử lý, không phải lỗi.")
            return None

    try:
        page.locator("tr.mat-row").first.wait_for(state="visible", timeout=12000)
        print("✅ Đã thấy danh sách văn bản.")
        text_utils.wait(page, 1000)
        return True
    except Exception as e:
        print("❌ Không thấy danh sách văn bản:", e)
        text_utils.save_debug(page, task.debug_prefix, "list_not_found")
        return False


def scroll_document_list_down(page, pixels: int = 900) -> bool:
    """
    Cuon dung vung danh sach van ban ben trai.

    DOffice/Angular khong cuon bang window, cung khong chac cuon bang mat-drawer.
    Danh sach nam trong 1 container an/ao, nen ham nay thu nhieu cach:
    1) dua chuot vao giua bang van ban roi wheel;
    2) tim moi phan tu dang scroll duoc o nua trai man hinh va cong scrollTop.
    """
    moved = False

    try:
        box = page.locator("xly-vbleft, table, tr.mat-row").first.bounding_box(timeout=3000)
        if box:
            x = box["x"] + min(box["width"] / 2, 260)
            y = box["y"] + min(max(box["height"] / 2, 250), 650)
            page.mouse.move(x, y)
            page.mouse.wheel(0, pixels)
            moved = True
    except Exception:
        pass

    try:
        js_moved = page.evaluate(
            """
            (pixels) => {
                const els = Array.from(document.querySelectorAll('*'));
                let moved = false;
                for (const el of els) {
                    const r = el.getBoundingClientRect();
                    const style = window.getComputedStyle(el);
                    const canScroll = el.scrollHeight > el.clientHeight + 20;
                    const inLeftDocPanel = r.left >= 120 && r.left < 700 && r.top < window.innerHeight - 80 && r.bottom > 120;
                    const overflowY = style.overflowY;
                    if (canScroll && inLeftDocPanel && ['auto','scroll','overlay','hidden'].includes(overflowY)) {
                        const old = el.scrollTop;
                        el.scrollTop = old + pixels;
                        if (el.scrollTop !== old) moved = true;
                    }
                }
                return moved;
            }
            """,
            pixels,
        )
        moved = moved or bool(js_moved)
    except Exception:
        pass

    text_utils.wait(page, 1200)
    return moved


def get_document_row(page, index_zero_based: int):
    """
    Lay row theo thu tu trong danh sach: 0 la van ban thu 1, 1 la van ban thu 2...
    Neu DOM chua load du row thi cuon dung container danh sach de Angular load them row.
    """
    rows = page.locator("tr.mat-row")

    for attempt in range(1, 13):
        count = rows.count()
        if count > index_zero_based:
            row = rows.nth(index_zero_based)
            try:
                row.scroll_into_view_if_needed(timeout=5000)
                row.wait_for(state="visible", timeout=5000)
                return row
            except Exception:
                pass

        print(
            f"⚠️ Chưa thấy row thứ {index_zero_based + 1}. DOM hiện có {count} row. "
            f"Thử cuộn vùng danh sách lần {attempt}/12..."
        )
        moved = scroll_document_list_down(page, pixels=900)
        if not moved:
            print("⚠️ Chưa xác định được container scroll; thử PageDown trên vùng danh sách...")
            try:
                page.keyboard.press("PageDown")
            except Exception:
                pass
            text_utils.wait(page, 1000)

    raise RuntimeError(f"Không lấy được văn bản thứ {index_zero_based + 1} trong danh sách.")


def click_document_row(row, prefer_flag_icon: bool = False) -> bool:
    print("\n--- Mở văn bản đang xử lý ---")

    click_candidates = []
    if prefer_flag_icon:
        # Chi danh sach "Chu tri - Da xu ly" co icon co (fa-icon) on dinh de click.
        click_candidates.append(
            (row.locator("div.vb-item > div").first.locator("fa-icon, .ng-fa-icon").first, "icon cờ trong đúng row")
        )

    click_candidates.extend(
        [
            (row.locator(".w-8").first, ".w-8 trong đúng row"),
            (row.locator("div.vb-item").first, "div.vb-item trong đúng row"),
            (row.locator("td.mat-cell").first, "td.mat-cell trong đúng row"),
        ]
    )

    for locator, desc in click_candidates:
        try:
            locator.scroll_into_view_if_needed(timeout=3000)
            locator.click(timeout=8000)
            print(f"✅ Đã click văn bản bằng {desc}.")
            text_utils.wait(row.page, 2500)
            return True
        except Exception as e:
            print(f"⚠️ Không click được {desc}: {e}")

    text_utils.save_debug(row.page, "doffice", "cannot_open_document_row")
    return False
