"""
Tim Chat ID Telegram de dien vao config.py (hoac trang web Cai dat) sau khi da
tao bot qua @BotFather.

CACH DUNG
---------
    1) Tao bot: mo Telegram, nhắn @BotFather -> /newbot -> lam theo huong dan
       -> nhan duoc 1 chuoi token dang "123456789:AAExxxxxxxxxxxxxxxxxxxxxxxxx".
    2) Nhan tin BAT KY (vd "hi") cho bot vua tao (tim theo username bot da dat),
       hoac neu muon nhan vao 1 nhom/group thi add bot vao nhom do va nhan 1
       tin bat ky trong nhom.
    3) Chay:
           python get_telegram_chat_id.py
       Script se hoi token (hoac tu doc tu config.py neu da dien san), goi
       Telegram de tim cac cuoc hoi thoai gan day va in ra Chat ID tuong ung.
    4) Dien Chat ID tim duoc vao config.py (TELEGRAM_CHAT_ID) hoac trang web
       Cai dat.

Neu khong thay gi, kiem tra lai da nhan tin cho DUNG bot chua, va thu lai (co
the mat vai giay de Telegram cap nhat).
"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

try:
    import config
except Exception:
    config = None


def get_token() -> str:
    default = getattr(config, "TELEGRAM_BOT_TOKEN", "") if config else ""
    prompt = f"Dán Bot Token (Enter để dùng token đã có trong config.py: {'có' if default else 'chưa có'}): "
    raw = input(prompt).strip()
    return raw or default


def main() -> None:
    print("=== Tìm Chat ID Telegram ===\n")
    token = get_token()
    if not token:
        print("❌ Chưa có Bot Token. Tạo bot qua @BotFather trước (xem docstring đầu file).")
        return

    url = f"https://api.telegram.org/bot{token}/getUpdates"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        print(f"❌ Lỗi HTTP {e.code} - kiểm tra lại Bot Token có đúng không.")
        return
    except Exception as e:
        print(f"❌ Không gọi được Telegram: {e}")
        return

    if not data.get("ok"):
        print("❌ Telegram báo lỗi:", data.get("description"))
        return

    results = data.get("result", [])
    if not results:
        print(
            "⚠️ Chưa thấy cuộc hội thoại nào.\n"
            "Nhắn tin (bất kỳ nội dung gì) cho bot trên Telegram trước, rồi chạy lại script này."
        )
        return

    seen = {}
    for update in results:
        msg = update.get("message") or update.get("channel_post") or {}
        chat = msg.get("chat")
        if not chat:
            continue
        chat_id = chat.get("id")
        title = chat.get("title") or chat.get("username") or chat.get("first_name") or "(không tên)"
        chat_type = chat.get("type", "")
        seen[chat_id] = f"{title}  [{chat_type}]"

    print("Các Chat ID tìm thấy:\n")
    for chat_id, label in seen.items():
        print(f"  {chat_id}   -   {label}")

    print("\nDán đúng Chat ID của bạn vào config.py (TELEGRAM_CHAT_ID) hoặc trang web Cài đặt.")


if __name__ == "__main__":
    main()
