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
from urllib.parse import urljoin, urlparse

from . import text_utils
from .task_types import TaskConfig


_STEP_TYPES = {"sidebar", "link_text", "link_href", "tab", "selector"}
_SELECTOR_ACTIONS = {"click", "wait"}


def _step_int(step: dict, name: str, default: int, minimum: int, maximum: int) -> int:
    """Doc mot so nguyen tu cau hinh web, giu trong khoang an toan."""
    try:
        value = int(step.get(name, default))
    except (TypeError, ValueError):
        value = default
    return max(minimum, min(maximum, value))


def _run_navigation_step(page, step: dict, task: TaskConfig, index: int) -> bool:
    """Chay mot buoc dieu huong khai bao trong TaskConfig.navigation_steps.

    Cac buoc "sidebar", "link_text", "link_href" va "tab" giu lai locator
    da duoc kiem chung cua du an. "selector" mo hoan toan cho Playwright
    locator (CSS, text=..., xpath=..., ...), de tung tai khoan co the chen them
    cac buoc trung gian ma khong sua code.
    """
    if not isinstance(step, dict):
        print(f"❌ Bước {index}: dữ liệu không hợp lệ (phải là object JSON).")
        return False

    step_type = str(step.get("type", "selector")).strip().lower()
    name = str(step.get("name") or f"Bước {index}").strip()
    optional = bool(step.get("optional", False))
    timeout = _step_int(step, "timeout_ms", 10000, 500, 60000)
    delay = _step_int(step, "delay_ms", 800, 0, 30000)

    if step_type not in _STEP_TYPES:
        message = f"Bước '{name}' có type không hỗ trợ: {step_type!r}."
        if optional:
            print(f"⚠️ {message} Bỏ qua vì optional=true.")
            return True
        print(f"❌ {message}")
        return False

    try:
        if step_type == "sidebar":
            value = str(step.get("value", "")).strip()
            if not value:
                raise ValueError("thiếu value")
            if not click_sidebar_item(page, value, task.debug_prefix):
                raise RuntimeError(f"không click được sidebar {value!r}")

        elif step_type == "link_text":
            value = str(step.get("value", "")).strip()
            if not value:
                raise ValueError("thiếu value")
            locator = page.get_by_role("link", name=re.compile(re.escape(value), re.I))
            if not text_utils.safe_click(locator, f"Link {value}", timeout=timeout):
                raise RuntimeError(f"không click được link {value!r}")

        elif step_type == "link_href":
            href = str(step.get("href") or step.get("value") or "").strip()
            if not href:
                raise ValueError("thiếu href")
            locator = page.locator(f'a[href="{href}"]').first
            if not text_utils.safe_click(locator, f"Link (href) {href}", timeout=timeout):
                raise RuntimeError(f"không click được href {href!r}")

        elif step_type == "tab":
            value = str(step.get("value", "")).strip()
            if not value:
                raise ValueError("thiếu value")
            if not click_tab(page, value):
                raise RuntimeError(f"không click được tab bắt buộc {value!r}")
            if step.get("empty_if_zero") and get_tab_count(page, value) == 0:
                print(f"ℹ️ Bước điều hướng {index}: tab {value} đang có 0 văn bản.")
                return "empty"

        else:  # selector
            selector = str(step.get("selector") or step.get("value") or "").strip()
            if not selector:
                raise ValueError("thiếu selector")
            action = str(step.get("action", "click")).strip().lower()
            if action not in _SELECTOR_ACTIONS:
                raise ValueError("action chỉ nhận 'click' hoặc 'wait'")
            nth = _step_int(step, "nth", 0, 0, 1000)
            locator = page.locator(selector).nth(nth)
            wait_for = str(step.get("wait_for", "visible")).strip().lower()
            if wait_for not in {"attached", "detached", "visible", "hidden"}:
                raise ValueError("wait_for chỉ nhận attached, detached, visible hoặc hidden")
            if action == "wait":
                locator.wait_for(state=wait_for, timeout=timeout)
            else:
                locator.click(timeout=timeout)

        print(f"✅ Bước điều hướng {index}: {name}.")
        if delay:
            text_utils.wait(page, delay)
        return True
    except Exception as e:
        if optional:
            print(f"⚠️ Bước điều hướng {index} '{name}' không chạy được, bỏ qua vì optional=true: {e}")
            return True
        print(f"❌ Bước điều hướng {index} '{name}' thất bại: {e}")
        return False


def run_navigation_steps(page, task: TaskConfig) -> bool:
    """Chay cac buoc dieu huong tuy chinh theo dung thu tu trong cau hinh."""
    steps = getattr(task, "navigation_steps", None) or []
    if not isinstance(steps, list):
        print("❌ navigation_steps phải là một mảng JSON.")
        return False
    for index, step in enumerate(steps, start=1):
        status = _run_navigation_step(page, step, task, index)
        if status == "empty":
            return "empty"
        if not status:
            return False
    return True


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


def click_tab(page, tab_name: str) -> bool:
    try:
        page.get_by_role("tab", name=re.compile(re.escape(tab_name), re.I)).click(timeout=10000)
        print(f"✅ Click tab {tab_name}.")
        return True
    except Exception as e:
        print(f"⚠️ Không click được tab {tab_name} bằng role, thử bằng text:", e)
        try:
            page.get_by_text(re.compile(re.escape(tab_name) + r"\s*\(\d+\)", re.I)).click(timeout=10000)
            print(f"✅ Click tab {tab_name} bằng text.")
            return True
        except Exception as e2:
            print(f"❌ Không click được tab bắt buộc {tab_name}:", e2)
            return False


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
    direct_path = str(getattr(task, "list_link_href", "") or "").strip()
    where = direct_path or f"{task.sidebar_item} / {task.list_link}"
    if task.tab_name:
        where += f" / {task.tab_name}"
    print(f"\n--- Vào {where} ---")

    page.goto(cfg.DOFFICE_URL, wait_until="domcontentloaded")
    text_utils.wait(page, 2500)

    click_login_if_needed(page)
    choose_role_if_needed(page, task.role_pattern, cfg.ROLE_BUTTON_NAME_HINT)

    # Che do don gian dung 3 truong sidebar/list/tab. JSON chi thay the luong
    # nay khi nguoi dung chu dong bat "Dieu huong nang cao". None giu tuong
    # thich voi task da tao tu ban truoc (co JSON thi xem la nang cao).
    advanced_navigation = getattr(task, "use_advanced_navigation", None)
    if advanced_navigation is None:
        advanced_navigation = bool(getattr(task, "navigation_steps", None))
    if advanced_navigation:
        if not getattr(task, "navigation_steps", None):
            print("❌ Đã bật điều hướng nâng cao nhưng chưa có bước JSON nào.")
            return False
        navigation_status = run_navigation_steps(page, task)
        if navigation_status == "empty":
            return None
        if not navigation_status:
            text_utils.save_debug(page, task.debug_prefix, "navigation_step_failed")
            return False
        try:
            row_selector = getattr(task, "document_row_selector", "tr.mat-row") or "tr.mat-row"
            page.locator(row_selector).first.wait_for(state="visible", timeout=12000)
            print("Da thay danh sach van ban qua navigation_steps.")
            text_utils.wait(page, 1000)
            return True
        except Exception as e:
            print("Khong thay danh sach van ban sau navigation_steps:", e)
            text_utils.save_debug(page, task.debug_prefix, "list_not_found")
            return False

    # URL truc tiep la cach on dinh nhat: sau khi chon Vai tro, vao thang route
    # danh sach va bo qua Sidebar/Tieu muc. Chi cho phep cung origin DOffice de
    # tranh cau hinh nham thanh mot website ben ngoai.
    if direct_path:
        direct_url = urljoin(cfg.DOFFICE_URL, direct_path)
        if urlparse(direct_url).netloc != urlparse(cfg.DOFFICE_URL).netloc:
            print(f"❌ Đường dẫn trực tiếp phải thuộc DOffice: {direct_path}")
            return False
        try:
            page.goto(direct_url, wait_until="domcontentloaded")
            text_utils.wait(page, 2000)
            print(f"✅ Đã mở đường dẫn trực tiếp: {direct_path}")
        except Exception as e:
            print(f"❌ Không mở được đường dẫn trực tiếp {direct_path}:", e)
            text_utils.save_debug(page, task.debug_prefix, "direct_link_failed")
            return False
    else:
        if not click_sidebar_item(page, task.sidebar_item, task.debug_prefix):
            return False
        text_utils.wait(page, 1000)
        link_locator = page.get_by_role("link", name=re.compile(re.escape(task.list_link), re.I))
        link_desc = f"Link {task.list_link}"
        if not text_utils.safe_click(link_locator, link_desc, timeout=10000):
            text_utils.save_debug(page, task.debug_prefix, "link_failed")
            return False
        text_utils.wait(page, 2000)

    if task.tab_name:
        if not click_tab(page, task.tab_name):
            text_utils.save_debug(page, task.debug_prefix, "tab_failed")
            return False
        text_utils.wait(page, 500)
        count = get_tab_count(page, task.tab_name)
        if count == 0:
            print(f"ℹ️ Tab {task.tab_name} hiện (0) văn bản - không còn gì cần xử lý, không phải lỗi.")
            return None

    try:
        row_selector = getattr(task, "document_row_selector", "tr.mat-row") or "tr.mat-row"
        page.locator(row_selector).first.wait_for(state="visible", timeout=12000)
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


def get_document_row(page, index_zero_based: int, row_selector: str = "tr.mat-row"):
    """
    Lay row theo thu tu trong danh sach: 0 la van ban thu 1, 1 la van ban thu 2...
    Neu DOM chua load du row thi cuon dung container danh sach de Angular load them row.
    """
    rows = page.locator(row_selector or "tr.mat-row")

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


def document_detail_is_open(page) -> bool:
    """Kiem tra click da mo chi tiet/PDF, khong chi focus dong danh sach."""
    # Dung cho test fake va khong anh huong Playwright that.
    if getattr(page, "document_detail_visible", False):
        return True

    candidates = [
        page.get_by_text(re.compile(r"File\s+văn bản", re.I)).first,
        page.get_by_role("button", name=re.compile(r"Tải xuống|Download", re.I)).first,
        page.locator("#download, button[title*='Tải xuống'], button[aria-label*='Download']").first,
        page.locator("iframe, embed, object, pdf-viewer, ngx-extended-pdf-viewer").first,
    ]
    for locator in candidates:
        try:
            locator.wait_for(state="visible", timeout=1200)
            return True
        except Exception:
            continue
    return False


def click_document_row(
    row, prefer_flag_icon: bool = False, extract_mode: str = "directive", document_click_selector: str = ""
) -> bool:
    print("\n--- Mở văn bản đang xử lý ---")

    click_candidates = []
    if document_click_selector:
        click_candidates.append((row.locator(document_click_selector).first, f"selector tuy chinh {document_click_selector}"))
    if extract_mode == "published":
        # Danh sach "Da phat hanh" (VB di) khong co div.vb-item/.w-8 (cau truc DOM
        # khac han van ban chi dao - xem extract.py). Toan bo td.mat-cell co click
        # handler chung (xac nhan qua test that: click bat ky span/div con nao ben
        # trong cell deu mo duoc van ban), nen click thang vao do, khong thu 2 cai
        # selector chac chan khong ton tai (do se timeout 3s x 2 = ton 6s vo ich).
        click_candidates.append((row.locator("td.mat-cell").first, "td.mat-cell trong đúng row (VB đã duyệt)"))
    else:
        try:
            has_vb_item = row.locator("div.vb-item").count() > 0
        except Exception:
            has_vb_item = False

        # Kieu thu ba: Cong viec ca nhan/ban phong (vd "Xem de biet") khong
        # co div.vb-item. Click handler nam tren chinh <tr mat-row>, click td co
        # the chi focus dong ma khong mo PDF. Click vao So VB o goc tren cua row
        # da duoc xac nhan mo dung detail/PDF cho ca KY_HIEU_CV va KY_HIEU_LDP.
        if not has_vb_item:
            click_candidates.extend(
                [
                    (
                        row.locator(
                            "td.mat-cell > div:first-child > span:first-child, "
                            "td > div:first-child > span:first-child"
                        ).first,
                        "Số văn bản trong tr.mat-row (Công việc)",
                    ),
                ]
            )
        elif prefer_flag_icon:
            # Chi danh sach "Chu tri - Da xu ly" co icon co (fa-icon) on dinh de click.
            click_candidates.append(
                (row.locator("div.vb-item > div").first.locator("fa-icon, .ng-fa-icon").first, "icon cờ trong đúng row")
            )

        if has_vb_item:
            click_candidates.extend(
                [
                    (row.locator(".w-8").first, ".w-8 trong đúng row"),
                    (row.locator("div.vb-item").first, "div.vb-item trong đúng row"),
                ]
            )

    for locator, desc in click_candidates:
        try:
            locator.scroll_into_view_if_needed(timeout=3000)
            locator.click(timeout=8000)
            # PDF viewer cua danh sach Xem de biet co the can gan 5 giay moi
            # render iframe; cho du truoc khi ket luan click chi focus dong.
            text_utils.wait(row.page, 4000)
            if document_detail_is_open(row.page):
                print(f"✅ Đã click văn bản bằng {desc}; chi tiết/PDF đã mở.")
                return True
            print(f"⚠️ Đã click {desc} nhưng chưa thấy chi tiết/PDF; thử selector kế tiếp.")
        except Exception as e:
            print(f"⚠️ Không click được {desc}: {e}")

    text_utils.save_debug(row.page, "doffice", "cannot_open_document_row")
    return False
