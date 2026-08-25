"""
Tien ich dung chung: "tee" stdout ra 1 file, van giu nguyen stdout that (con
console/terminal goc, hoac 1 tee khac dang bao boc no - vi du webapp/run_manager.py
boc stdout de phat nhat ky truc tiep len web). Nho tinh chat long nhau nay, du
chay bang CLI, Scheduled Task, hay tu web, moi lan goi runner.run_selected_tasks()
deu tu dong co 1 file log rieng, khong can lam gi them o noi goi.
"""
from __future__ import annotations

import io


class FileTeeStream(io.TextIOBase):
    def __init__(self, original, file_handle) -> None:
        self._original = original
        self._file = file_handle

    def write(self, s: str) -> int:
        # Khi chay bang pythonw.exe, Windows khong tao console va sys.stdout
        # bang None. File log van la dich ghi hop le va phai tiep tuc hoat dong.
        if self._original is not None:
            self._original.write(s)
        self._file.write(s)
        self._file.flush()
        return len(s)

    def flush(self) -> None:
        if self._original is not None:
            self._original.flush()
        self._file.flush()
