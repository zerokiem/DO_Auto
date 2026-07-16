"""
Trich xuat thong tin van ban tu 1 dong (row) trong danh sach DOffice.

Gop lai tu ban trich xuat cua DO_phoi_hop.py va DO_Dang_Doan_phoi_hop_v3.py (ca
hai deu dung page.evaluate JS vi HTML cua DOffice o cac man hinh nay hoi khac
nhau), them ca cach lay khoi chi dao dang <lib-view-noi-dung-thuc-hien> (rieng
cua man hinh Cong viec Chi bo/Cong doan) lan <section class="text-blue-600">
(cac man hinh Van ban chuyen mon), de 1 ham duy nhat dung on dinh cho ca 3 man
hinh danh sach.
"""
from __future__ import annotations

import re
from typing import Dict

from . import text_utils

_ROW_EVAL_JS = r"""
(row) => {
    const clean = (s) => (s || '')
        .replace(/\u00a0/g, ' ')
        .replace(/[ \t]+/g, ' ')
        .replace(/\n\s+/g, '\n')
        .trim();

    const vb = row.querySelector('div.vb-item');
    if (!vb) return {};

    const directChildren = Array.from(vb.children);
    const directDivs = directChildren.filter(el => el.tagName && el.tagName.toLowerCase() === 'div');

    const firstDiv = directDivs[0] || null;
    const firstText = clean(firstDiv ? firstDiv.innerText : '');

    let soVb = '';
    if (firstDiv) {
        const firstSpan = firstDiv.querySelector(':scope > span');
        soVb = clean(firstSpan ? firstSpan.innerText : '');
        if (!soVb && firstText) soVb = clean(firstText.split('\n')[0]);
    }

    let ngayVb = '';
    const dateMatch = firstText.match(/\d{2}\/\d{2}\/\d{4}/);
    if (dateMatch) ngayVb = dateMatch[0];

    const secondDiv = directDivs[1] || null;
    let noiPhatHanh = '';
    if (secondDiv) {
        const blue = secondDiv.querySelector('span.text-blue-600, span[style*="color"]');
        const firstSpan = secondDiv.querySelector(':scope > span');
        noiPhatHanh = clean((blue || firstSpan || secondDiv).innerText);
        noiPhatHanh = clean(noiPhatHanh.split('\n')[0]);
    }

    // Trich yeu: thu nhieu cach vi man hinh Chu tri/Phoi hop/Dang doan render hoi khac nhau.
    let trichYeu = '';
    const thirdDiv = directDivs[2] || null;
    if (thirdDiv) {
        const spans = Array.from(thirdDiv.children).filter(el => el.tagName && el.tagName.toLowerCase() === 'span');
        const summarySpan = spans.find(el => !/^Chủ trì:|^Phối hợp:/i.test(clean(el.innerText)));
        if (summarySpan) trichYeu = clean(summarySpan.innerText);
    }
    if (!trichYeu) {
        const directSpan = directChildren.find(el => el.tagName && el.tagName.toLowerCase() === 'span');
        if (directSpan) trichYeu = clean(directSpan.innerText);
    }
    if (!trichYeu) {
        const unread = vb.querySelector('span.unread');
        trichYeu = clean(unread ? unread.innerText : '');
    }
    if (!trichYeu && thirdDiv) {
        trichYeu = clean(thirdDiv.innerText.split('\n')[0]);
    }

    // Khoi chi dao: thu nhieu cau truc (van ban chuyen mon, cong viec Chi bo/Cong doan).
    const chiDaoEl = vb.querySelector('lib-view-noi-dung-thuc-hien')
        || vb.querySelector('section.text-blue-600 section')
        || vb.querySelector('section.text-blue-600')
        || null;
    const chiDaoText = clean(chiDaoEl ? chiDaoEl.innerText : '');

    return {
        so_vb: soVb,
        ngay_vb: ngayVb,
        noi_phat_hanh: noiPhatHanh,
        trich_yeu: trichYeu,
        chi_dao_text: chiDaoText,
        row_text: clean(vb.innerText),
    };
}
"""

_EMPTY_DATA = {
    "so_vb": "",
    "ngay_vb": "",
    "noi_phat_hanh": "",
    "trich_yeu": "",
    "nguoi_chi_dao": "",
    "thoi_gian_chi_dao": "",
    "noi_dung_chi_dao": "",
    "chu_tri": "",
    "phoi_hop": "",
}


def extract_document_info_from_row(row) -> Dict[str, str]:
    data = dict(_EMPTY_DATA)

    try:
        raw = row.evaluate(_ROW_EVAL_JS) or {}
    except Exception as e:
        print(f"⚠️ Không evaluate được row để trích metadata: {e}")
        raw = {}

    data["so_vb"] = text_utils.clean_text(raw.get("so_vb", ""))
    data["ngay_vb"] = text_utils.clean_date_vb(raw.get("ngay_vb", ""))
    data["noi_phat_hanh"] = text_utils.clean_text(raw.get("noi_phat_hanh", ""))
    data["trich_yeu"] = text_utils.clean_text(raw.get("trich_yeu", ""))
    data.update(text_utils.parse_chi_dao_block(raw.get("chi_dao_text", "")))

    # Fallback cuoi neu DOM render khac hoan toan cau truc mong doi o tren.
    if not data["so_vb"] or not data["ngay_vb"] or not data["trich_yeu"]:
        row_text = text_utils.clean_text(raw.get("row_text", ""))
        lines = [text_utils.clean_text(x) for x in row_text.splitlines() if text_utils.clean_text(x)]

        if not data["so_vb"] and lines:
            data["so_vb"] = lines[0]

        if not data["ngay_vb"]:
            m = re.search(r"\d{2}/\d{2}/\d{4}", row_text)
            if m:
                data["ngay_vb"] = m.group(0)

        if not data["noi_phat_hanh"] and len(lines) >= 2:
            for line in lines[1:5]:
                if re.fullmatch(r"\d{2}/\d{2}/\d{4}", line):
                    continue
                if line.lower().startswith(("chủ trì:", "phối hợp:")):
                    continue
                if re.search(r"\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}", line):
                    continue
                if line != data["so_vb"]:
                    data["noi_phat_hanh"] = line
                    break

        if not data["trich_yeu"]:
            for line in lines:
                lower = line.lower()
                if line in (data["so_vb"], data["ngay_vb"], data["noi_phat_hanh"]):
                    continue
                if lower.startswith(("chủ trì:", "phối hợp:")):
                    continue
                if re.search(r"\d{2}/\d{2}/\d{4}", line):
                    continue
                data["trich_yeu"] = line
                break

    return data


def print_document_info(data: Dict[str, str]) -> None:
    print("\n📄 THÔNG TIN VĂN BẢN")
    print(f"  - Số VB:              {data.get('so_vb', '')}")
    print(f"  - Ngày VB:            {data.get('ngay_vb', '')}")
    print(f"  - Nơi phát hành:      {data.get('noi_phat_hanh', '')}")
    print(f"  - Trích yếu:          {data.get('trich_yeu', '')}")
    print(f"  - Người chỉ đạo:      {data.get('nguoi_chi_dao', '')}")
    print(f"  - Thời gian chỉ đạo:  {data.get('thoi_gian_chi_dao', '')}")
    print(f"  - Nội dung chỉ đạo:   {data.get('noi_dung_chi_dao', '')}")
    print(f"  - Chủ trì:            {data.get('chu_tri', '')}")
    print(f"  - Phối hợp:           {data.get('phoi_hop', '')}")
