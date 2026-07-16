"""
Gui tin nhan tong ket phien chay qua Telegram Bot API.

Vi sao chon Telegram: mien phi hoan toan, khong gioi han, API cong khai don
gian (1 request HTTP la gui duoc tin nhan), khong can xac minh doanh nghiep
nhu Zalo Official Account hay WhatsApp Business API chinh chu (2 dich vu nay
kho dung cho 1 nguoi/noi bo, phai qua duyet). Cac thu vien WhatsApp/Zalo "khong
chinh thuc" (dang nhap bang QR quet phien ca nhan) de vi pham dieu khoan dich
vu va de bi khoa tai khoan, nen khong dung o day.

Xem README de biet cach tao bot (@BotFather) va lay chat_id
(get_telegram_chat_id.py).
"""
from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import List

from .task_types import TaskResult

TELEGRAM_API_BASE = "https://api.telegram.org"
MAX_DOCS_LISTED_PER_TASK = 10
MAX_MESSAGE_LENGTH = 4000  # Telegram gioi han 4096 ky tu/tin nhan, chua an toan


def _escape_html(text: str) -> str:
    return (text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_summary_message(results: List[TaskResult], trigger_source: str, test_mode: bool = False) -> str:
    tag = " [CHẾ ĐỘ TEST]" if test_mode else ""
    lines = [f"📋 <b>DOffice - Tổng kết phiên chạy</b>{tag}", f"🕐 Nguồn: {_escape_html(trigger_source)}", ""]

    total_new = 0
    for r in results:
        icon = "✅" if r.ok else "❌"
        header = f"{icon} <b>{_escape_html(r.label)}</b>"

        if not r.ok:
            header += f" — Lỗi: {_escape_html(r.note or 'không rõ')}"
            lines.append(header)
            lines.append("")
            continue

        header += f": {r.processed} văn bản mới"
        lines.append(header)
        total_new += r.processed

        for doc in (r.documents or [])[:MAX_DOCS_LISTED_PER_TASK]:
            so_vb = _escape_html(doc.get("so_vb", ""))
            trich_yeu = _escape_html(doc.get("trich_yeu", ""))
            if len(trich_yeu) > 70:
                trich_yeu = trich_yeu[:70] + "…"
            lines.append(f"  • {so_vb} - {trich_yeu}" if trich_yeu else f"  • {so_vb}")

        if len(r.documents or []) > MAX_DOCS_LISTED_PER_TASK:
            lines.append(f"  … và {len(r.documents) - MAX_DOCS_LISTED_PER_TASK} văn bản khác")

        lines.append("")

    lines.append(f"<b>Tổng: {total_new} văn bản mới trong phiên này.</b>")

    message = "\n".join(lines)
    if len(message) > MAX_MESSAGE_LENGTH:
        message = message[:MAX_MESSAGE_LENGTH] + "\n\n… (đã cắt bớt, xem đầy đủ trong Excel/Lịch sử)"
    return message


def send_telegram_message(bot_token: str, chat_id: str, text: str) -> "tuple[bool, str]":
    """Gui 1 tin nhan qua Telegram Bot API. Tra ve (thanh_cong, thong_bao_loi_neu_co)."""
    if not bot_token or not chat_id:
        return False, "Chưa cấu hình TELEGRAM_BOT_TOKEN hoặc TELEGRAM_CHAT_ID."

    url = f"{TELEGRAM_API_BASE}/bot{bot_token}/sendMessage"
    payload = json.dumps(
        {"chat_id": chat_id, "text": text, "parse_mode": "HTML", "disable_web_page_preview": True}
    ).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"}, method="POST")

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            if body.get("ok"):
                return True, ""
            return False, body.get("description", "Không rõ lỗi.")
    except urllib.error.HTTPError as e:
        try:
            detail = json.loads(e.read().decode("utf-8")).get("description", str(e))
        except Exception:
            detail = str(e)
        return False, f"HTTP {e.code}: {detail}"
    except urllib.error.URLError as e:
        return False, f"Không kết nối được Telegram: {e.reason}"
    except Exception as e:
        return False, str(e)


def notify_run_summary(cfg, results: List[TaskResult], trigger_source: str, test_mode: bool = False) -> None:
    """Goi sau khi 1 phien chay hoan tat. Khong bao gio raise loi ra ngoai - neu
    gui that bai chi in canh bao, khong lam hong ket qua tu dong hoa da chay
    xong. Bo qua hoan toan neu ENABLE_TELEGRAM_NOTIFY=False (mac dinh)."""
    if not getattr(cfg, "ENABLE_TELEGRAM_NOTIFY", False):
        return

    try:
        total_new = sum(r.processed for r in results if r.ok)
        if getattr(cfg, "TELEGRAM_NOTIFY_ONLY_IF_NEW", False) and total_new == 0:
            print("⏸️ Bỏ qua gửi Telegram vì không có văn bản mới (TELEGRAM_NOTIFY_ONLY_IF_NEW=True).")
            return

        message = build_summary_message(results, trigger_source, test_mode=test_mode)
        ok, error = send_telegram_message(cfg.TELEGRAM_BOT_TOKEN, cfg.TELEGRAM_CHAT_ID, message)
        if ok:
            print("✅ Đã gửi tin nhắn tổng kết qua Telegram.")
        else:
            print(f"⚠️ Không gửi được tin nhắn tổng kết qua Telegram: {error}")
    except Exception as e:
        print(f"⚠️ Lỗi không mong muốn khi gửi Telegram: {e}")
