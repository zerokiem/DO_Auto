"""
Khoi dong web dashboard cho DOffice.

CACH DUNG
---------
    python run_web.py

Mac dinh mo tai http://127.0.0.1:8877 va tu dong mo trinh duyet mac dinh sau ~1
giay. Dong cua so terminal (hoac Ctrl+C) de tat server.

TRUY CAP TU XA (dien thoai/laptop khac qua Tailscale)
------------------------------------------------------
Doi HOST ben duoi thanh "0.0.0.0" roi chay lai, sau do vao dia chi Tailscale
cua may nay tu thiet bi khac, vi du: http://100.x.y.z:8877

LUU Y AN TOAN: day la cong cu noi bo dung 1 minh (hoac vai nguoi tin cay qua
Tailscale), KHONG co man hinh dang nhap rieng. Khong mo port nay ra internet
cong khai (khong port-forward tren router).

LUU Y VE TRINH DUYET TU DONG: khi chon "Chay hien cua so Chromium" (bo tick o
o "Chay an"), cua so Chromium se mo ngay TREN MAY DANG CHAY SERVER NAY, khong
phai tren may ban dang xem trinh duyet dashboard. Neu kich hoat tu dien thoai
qua Tailscale, nen de "Chay an" (headless) vi se khong co ai o do de xem cua so.
"""
from __future__ import annotations

import threading
import time
import webbrowser

from webapp.app import app

HOST = "127.0.0.1"  # doi thanh "0.0.0.0" de truy cap tu may khac qua Tailscale/LAN
PORT = 8877


def _open_browser_later() -> None:
    time.sleep(1.2)
    try:
        webbrowser.open(f"http://127.0.0.1:{PORT}")
    except Exception:
        pass


if __name__ == "__main__":
    threading.Thread(target=_open_browser_later, daemon=True).start()
    print(f"DOffice web dashboard: http://{HOST}:{PORT}   (Ctrl+C để dừng)")
    app.run(host=HOST, port=PORT, threaded=True, debug=False)
