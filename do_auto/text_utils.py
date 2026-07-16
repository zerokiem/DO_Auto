"""
Ham phu dung chung: don dep text, tao ten file an toan, parse khoi chi dao.

Gop tu 3 script goc (DO_chu_tri_da_XL_v3.py, DO_phoi_hop.py,
DO_Dang_Doan_phoi_hop_v3.py) - logic 3 file nay giong het nhau nen chi giu 1 ban.
"""
from __future__ import annotations

import re
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Dict


def wait(page, ms: int = 1000) -> None:
    page.wait_for_timeout(ms)


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s+", "\n", text)
    return text.strip()


def clean_date_vb(text: str) -> str:
    text = clean_text(text)
    m = re.search(r"\d{2}/\d{2}/\d{4}", text)
    return m.group(0) if m else text


def normalize_key(text: str) -> str:
    text = clean_text(text or "")
    text = text.lower()
    text = re.sub(r"\s+", "", text)
    return text


def parse_chi_dao_block(block_text: str) -> Dict[str, str]:
    """
    Dang thuong gap:
    Nguyen Xuan Binh - 21/05/2026 15:12:00
    Noi dung chi dao...
    Chu tri: TKT
    Phoi hop: TDDDXO
    """
    result = {
        "nguoi_chi_dao": "",
        "thoi_gian_chi_dao": "",
        "noi_dung_chi_dao": "",
        "chu_tri": "",
        "phoi_hop": "",
    }

    block_text = clean_text(block_text)
    if not block_text:
        return result

    lines = [clean_text(x) for x in block_text.splitlines() if clean_text(x)]
    if not lines:
        return result

    first = lines[0]
    m = re.match(r"^(.*?)\s*-\s*(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2})$", first)
    if m:
        result["nguoi_chi_dao"] = clean_text(m.group(1))
        result["thoi_gian_chi_dao"] = clean_text(m.group(2))
    else:
        result["nguoi_chi_dao"] = first

    noi_dung_lines = []
    for line in lines[1:]:
        lower = line.lower()
        if lower.startswith("chủ trì:"):
            result["chu_tri"] = clean_text(line.split(":", 1)[1])
        elif lower.startswith("phối hợp:"):
            result["phoi_hop"] = clean_text(line.split(":", 1)[1])
        else:
            noi_dung_lines.append(line)

    result["noi_dung_chi_dao"] = "\n".join(noi_dung_lines)
    return result


def remove_vietnamese_accents(text: str) -> str:
    """Bo dau tieng Viet, chuyen D/d sang D/d thuong."""
    text = text.replace("Đ", "D").replace("đ", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", text)


def safe_ascii_filename(name: str) -> str:
    """
    Lam sach ten file:
    - bo dau tieng Viet;
    - bo ky tu dac biet;
    - chi giu chu/so/khoang trang/gach duoi/gach ngang/dau cham.
    """
    name = remove_vietnamese_accents(name or "")
    name = name.replace("/", " ").replace("\\", " ")
    name = re.sub(r"[^A-Za-z0-9 ._\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" ._-")
    return name or f"van_ban_{int(time.time())}"


def extract_so_vb_code(so_vb: str) -> str:
    """
    Lay phan so trong So VB de dua vao dau ngu ten file.

    Vi du:
    - 820/TTĐĐN     -> 0820
    - 172/TB-CĐPTC4 -> 0172
    - 2190/PTC4-KT  -> 2190
    - 12345/ABC     -> 12345

    Neu khong tim thay so thi dung 0000 de chuong trinh van chay tiep.
    """
    so_vb = clean_text(so_vb or "")

    m = re.match(r"^\s*(\d+)", so_vb)
    if not m:
        m = re.search(r"\d+", so_vb)

    if not m:
        return "0000"

    digits = m.group(0)
    return digits.zfill(4) if len(digits) < 4 else digits


def add_date_prefix_if_needed(filename: str, so_vb: str) -> str:
    """
    Them dau ngu yymmdd-xxxx - vao ten file neu chua co.
    xxxx lay tu phan so cua So VB, it hon 4 chu so thi them so 0 phia truoc.
    """
    if re.match(r"^\d{6}-\d{4,}\s+-\s+", filename):
        return filename

    prefix = datetime.now().strftime("%y%m%d")
    vb_code = extract_so_vb_code(so_vb)
    return f"{prefix}-{vb_code} - {filename}"


def unique_path(folder: Path, filename: str) -> Path:
    filename = safe_ascii_filename(filename)
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    path = folder / filename
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    for i in range(1, 9999):
        new_path = folder / f"{stem}_{i}{suffix}"
        if not new_path.exists():
            return new_path

    raise RuntimeError("Không tạo được tên file không trùng.")


def save_debug(page, prefix: str, name: str) -> None:
    path = f"debug_{prefix}_{name}.png"
    try:
        page.screenshot(path=path, full_page=True)
        print(f"📸 Đã lưu ảnh debug: {path}")
    except Exception as e:
        print(f"⚠️ Không chụp được ảnh debug: {e}")


def safe_click(locator, desc: str, timeout: int = 8000) -> bool:
    try:
        locator.wait_for(state="visible", timeout=timeout)
        locator.click(timeout=timeout)
        print(f"✅ Click: {desc}")
        return True
    except Exception as e:
        print(f"❌ Không click được {desc}: {e}")
        return False
