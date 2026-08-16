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

    // Mot so trang Cong viec (vd "Xem de biet") khong co div.vb-item ma
    // dat toan bo metadata truc tiep trong td.mat-cell. Day la mot cau truc
    // rieng: 4 truong metadata phai duoc lay bang selector tuong doi trong row,
    // khong dung innerText cua ca cell (se tron So VB/Ngay/Noi PH/Trich yeu).
    const workCell = row.querySelector(
        'td.mat-column-KY_HIEU_CV, td.mat-column-KY_HIEU_LDP, td.mat-column-KY_HIEU_LDDV'
    );
    const isWorkList = workCell && !workCell.querySelector('div.vb-item');
    const vb = row.querySelector('div.vb-item') || workCell || row.querySelector('td.mat-cell') || row;

    const directChildren = Array.from(vb.children);
    const directDivs = directChildren.filter(el => el.tagName && el.tagName.toLowerCase() === 'div');

    const firstDiv = directDivs[0] || null;
    const firstText = clean(firstDiv ? firstDiv.innerText : '');

    let soVb = '';
    if (isWorkList && firstDiv) {
        // Xem de biet (giao dien moi): td > div[1] > span[1].
        // :nth-of-type giu dung y nghia khi Angular chen comment node.
        const firstSpan = firstDiv.querySelector(':scope > span:nth-of-type(1)');
        soVb = clean(firstSpan ? firstSpan.innerText : '');
    }
    if (!soVb && firstDiv) {
        const firstSpan = firstDiv.querySelector(':scope > span');
        soVb = clean(firstSpan ? firstSpan.innerText : '');
        if (!soVb && firstText) soVb = clean(firstText.split('\n')[0]);
    }

    let ngayVb = '';
    if (isWorkList && firstDiv) {
        // DOffice dang co 2 bien the: span[2] la sibling cua So VB, hoac nam
        // trong div.dokhan. Ca hai deu nam trong div metadata thu nhat.
        const dateEl = firstDiv.querySelector(
            ':scope > span:nth-of-type(2), :scope > span.mat-tooltip-trigger, :scope > .dokhan > span'
        );
        ngayVb = clean(dateEl ? dateEl.innerText : '');
    }
    if (!ngayVb) {
        const dateMatch = firstText.match(/\d{2}\/\d{2}\/\d{4}/);
        if (dateMatch) ngayVb = dateMatch[0];
    }

    const secondDiv = directDivs[1] || null;
    let noiPhatHanh = '';
    if (secondDiv) {
        // Xem de biet: td > div[2] > span.text-blue-600 (giao dien moi),
        // giao dien cu dung inline style color. Chi lay span truc tiep de
        // khong lay nham cac dong noi dung/chi dao ben duoi.
        const blue = secondDiv.querySelector(
            ':scope > span.text-blue-600, :scope > span[style*="color"]'
        );
        const firstSpan = secondDiv.querySelector(':scope > span:nth-of-type(1)');
        noiPhatHanh = clean((blue || firstSpan || secondDiv).innerText);
        noiPhatHanh = clean(noiPhatHanh.split('\n')[0]);
    }

    // Trich yeu: thu nhieu cach vi man hinh Chu tri/Phoi hop/Dang doan render hoi khac nhau.
    let trichYeu = '';
    const thirdDiv = directDivs[2] || null;
    if (isWorkList) {
        // Bien the duoc inspect: td > span la Trich yeu. Bien the dang dung
        // tren tai khoan kiem thu: td > div[3] > span[1]. Ho tro ca hai.
        const directSummary = vb.querySelector(':scope > span');
        const nestedSummary = thirdDiv && thirdDiv.querySelector(':scope > span:nth-of-type(1)');
        trichYeu = clean((directSummary || nestedSummary) ? (directSummary || nestedSummary).innerText : '');
    }
    if (!trichYeu && thirdDiv) {
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

    // Fallback rieng cho Xem de biet neu DOffice doi class/span: theo dung
    // thu tu hien thi So VB -> Ngay VB -> Noi phat hanh -> Ngay nhan -> Trich yeu.
    // Khong ap dung cho cac list khac de tranh tach nham noi dung chi dao.
    if (isWorkList) {
        const workLayout = clean(vb.innerText).match(
            /^(.+?)\s+(\d{2}\/\d{2}\/\d{4})\s+(.+?)\s+\d{2}\/\d{2}\/\d{4}\s+(.+)$/
        );
        if (workLayout) {
            const [, regexSoVb, regexNgayVb, regexNoiPhatHanh, regexTrichYeu] = workLayout;
            if (!soVb || /\d{2}\/\d{2}\/\d{4}/.test(soVb)) soVb = clean(regexSoVb);
            if (!ngayVb) ngayVb = clean(regexNgayVb);
            if (!noiPhatHanh || /\d{2}\/\d{2}\/\d{4}/.test(noiPhatHanh)) noiPhatHanh = clean(regexNoiPhatHanh);
            if (!trichYeu || /\d{2}\/\d{2}\/\d{4}/.test(trichYeu)) trichYeu = clean(regexTrichYeu);
        }
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

# Trich cho VAN BAN DA KY DUYET PHAT HANH (VB di) - cau truc DOM khac han van ban
# chi dao: KHONG co div.vb-item. Trong td.mat-cell:
#   dong 1: <span style="font-weight:bold">So VB</span> ... <span class="time">Ngay phat hanh</span>
#   dong 2: <div class="d-flex"><span>Nguoi soan thao - </span><span>Don vi soan thao</span></div>
#   dong 3: <div>Trich yeu</div>
_PUBLISHED_ROW_EVAL_JS = r"""
(row) => {
    const clean = (s) => (s || '')
        .replace(/[\s\u00a0]+/g, ' ')
        .replace(/[ \t]+/g, ' ')
        .replace(/\n\s+/g, '\n')
        .trim();

    const cell = row.querySelector('td.mat-cell') || row;

    let soVb = '';
    const boldSpan = cell.querySelector('span[style*="font-weight: bold"], span[style*="font-weight:bold"]');
    if (boldSpan) soVb = clean(boldSpan.innerText);

    let ngayRaw = '';
    const timeSpan = cell.querySelector('span.time, span[title="Ngày phát hành"]');
    if (timeSpan) ngayRaw = clean(timeSpan.innerText);

    let nguoi = '', donVi = '';
    const dflex = cell.querySelector('div.d-flex');
    if (dflex) {
        const spans = Array.from(dflex.querySelectorAll(':scope > span'));
        if (spans.length >= 2) {
            nguoi = clean(spans[0].innerText).replace(/[-–\s]+$/, '').trim();
            donVi = clean(spans[1].innerText);
        } else {
            const parts = clean(dflex.innerText).split(/\s+-\s+/);
            if (parts.length >= 2) {
                nguoi = clean(parts[0]);
                donVi = clean(parts.slice(1).join(' - '));
            } else {
                nguoi = clean(dflex.innerText);
            }
        }
    }

    let trichYeu = '';
    if (dflex) {
        let sib = dflex.nextElementSibling;
        while (sib) {
            const st = sib.getAttribute('style') || '';
            const txt = clean(sib.innerText);
            const hasIcon = sib.querySelector('fa-icon, svg');
            const isTomato = /tomato/i.test(st);
            if (txt && !hasIcon && !isTomato) { trichYeu = txt; break; }
            sib = sib.nextElementSibling;
        }
    }

    return {
        so_vb: soVb,
        ngay_raw: ngayRaw,
        nguoi_soan_thao: nguoi,
        don_vi_soan_thao: donVi,
        trich_yeu: trichYeu,
        row_text: clean(cell.innerText),
    };
}
"""


def extract_published_doc_info_from_row(row) -> Dict[str, str]:
    """Trich thong tin van ban DA KY DUYET PHAT HANH (VB di) tu 1 row.

    Map vao dung cot Excel san co (theo lua chon nguoi dung - dung lai 13 cot):
      - Don vi soan thao -> cot 'Noi phat hanh' (noi_phat_hanh)
      - Nguoi soan thao  -> cot 'Nguoi chi dao' (nguoi_chi_dao)
    Cac cot chi dao con lai de trong."""
    data = dict(_EMPTY_DATA)
    try:
        raw = row.evaluate(_PUBLISHED_ROW_EVAL_JS) or {}
    except Exception as e:
        print(f"⚠️ Không evaluate được row (VB đã duyệt) để trích metadata: {e}")
        raw = {}

    data["so_vb"] = text_utils.clean_text(raw.get("so_vb", ""))
    data["ngay_vb"] = text_utils.clean_date_vb(raw.get("ngay_raw", ""))
    data["noi_phat_hanh"] = text_utils.clean_text(raw.get("don_vi_soan_thao", ""))
    data["nguoi_chi_dao"] = text_utils.clean_text(raw.get("nguoi_soan_thao", ""))
    data["trich_yeu"] = text_utils.clean_text(raw.get("trich_yeu", ""))

    # Fallback neu DOffice doi cau truc DOM: lay tu toan bo text cua row.
    if not data["so_vb"] or not data["ngay_vb"] or not data["trich_yeu"]:
        row_text = text_utils.clean_text(raw.get("row_text", ""))
        lines = [text_utils.clean_text(x) for x in row_text.splitlines() if text_utils.clean_text(x)]
        if not data["so_vb"] and lines:
            data["so_vb"] = lines[0]
        if not data["ngay_vb"]:
            m = re.search(r"\d{2}/\d{2}/\d{4}", row_text)
            if m:
                data["ngay_vb"] = m.group(0)
        if not data["trich_yeu"] and lines:
            data["trich_yeu"] = max(lines, key=len)  # dong dai nhat thuong la trich yeu

    return data


def _is_work_list_row(row) -> bool:
    """Nhan biet hang Cong viec/Xem de biet truoc khi chon parser.

    Nguoi dung co the da chon nham "Van ban da phat hanh" tren trang Cai dat.
    Hai cot Angular KY_HIEU_CV/KY_HIEU_LDP la dau hieu DOM on dinh hon ten task,
    nen uu tien cau truc that cua hang de tranh fallback lay ca o thanh So VB.
    """
    try:
        return bool(
            row.evaluate(
                "(row) => Boolean(row.querySelector('td.mat-column-KY_HIEU_CV, td.mat-column-KY_HIEU_LDP, td.mat-column-KY_HIEU_LDDV'))"
            )
        )
    except Exception:
        return False


def extract_document_info_from_row(row, mode: str = "directive") -> Dict[str, str]:
    is_work_list = _is_work_list_row(row)
    if mode == "published" and not is_work_list:
        return extract_published_doc_info_from_row(row)
    if mode == "published" and is_work_list:
        print("ℹ️ Phát hiện dòng Công việc/Xem để biết; dùng parser metadata riêng thay cho dạng đã phát hành.")

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
