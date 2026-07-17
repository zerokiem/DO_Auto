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

import sys
import threading
import time
import webbrowser

from do_auto.text_utils import fix_windows_console_encoding

fix_windows_console_encoding()

import config as base_config
from webapp.app import app, run_manager

#HOST = "127.0.0.1"  # doi thanh "0.0.0.0" de truy cap tu may khac qua Tailscale/LAN
HOST = "0.0.0.0"
PORT = 8877


def _open_browser_later() -> None:
    time.sleep(1.2)
    try:
        webbrowser.open(f"http://127.0.0.1:{PORT}")
    except Exception:
        pass


if __name__ == "__main__":
    if sys.platform == "win32":
        # Tren NAS/Docker (khong man hinh) khong co trinh duyet nao de mo, va
        # tab "Lich chay" tren Linux dung thread nen rieng (khong phai Windows
        # Task Scheduler) - xem do_auto/inprocess_scheduler.py.
        threading.Thread(target=_open_browser_later, daemon=True).start()
    else:
        from do_auto import inprocess_scheduler

        inprocess_scheduler.start_background_scheduler(run_manager, base_config)

    print(f"DOffice web dashboard: http://{HOST}:{PORT}   (Ctrl+C để dừng)")
    app.run(host=HOST, port=PORT, threaded=True, debug=False)
