# DOffice (bản gộp + web dashboard) – README

Bộ công cụ tự động tải PDF, tổng hợp Excel và xử lý văn bản trên DOffice bằng
Python + Playwright. Có 2 cách dùng: **CLI/PowerShell** (như trước) hoặc **bảng
điều khiển web** chạy ngay trên máy bạn (mục 7).

> **Đây là hướng dẫn chạy trực tiếp trên Windows** (máy cá nhân/máy cơ quan,
> Playwright/Chromium chạy ngay trên máy đó, không cần Docker). Nếu bạn muốn
> chạy 24/7 trên NAS Synology qua Docker (không cần mở máy tính), xem
> [`README_NAS.md`](README_NAS.md) – cùng 1 bộ code, chỉ khác cách triển khai
> (biến môi trường `DOFFICE_DATA_DIR`/`DOFFICE_DISPLAY_URL`
> tự động chuyển hành vi, xem mục 10).

| Trước đây (3 file rời) | Bây giờ |
|---|---|
| `DO_chu_tri_da_XL_v3.py` | tác vụ `chu_tri` trong `config.py` |
| `DO_phoi_hop.py` | tác vụ `phoi_hop` |
| `DO_Dang_Doan_phoi_hop_v3.py` | tác vụ `dang_doan` |
| Các file Excel riêng | **1 file Excel `Tong_hop_DOffice.xlsx`, mỗi tác vụ có một sheet riêng** |
| Vai trò "Phó Truyền tải điện" hard-code trong code | Gõ vai trò trong `config.py` hoặc trang web **Cài đặt** (ghi thẳng vào `config.py`) |
| Chạy `python DO_xxx.py` từng file | `python run_doffice.py` (menu/`--tasks`/`--all`) **hoặc** `python run_web.py` (bảng điều khiển web) |
| Chỉ có log khi chạy Scheduled Task | **Mọi lần chạy** (tay, Scheduled Task, web) đều tự tạo 1 file log riêng |
| Không có lịch sử chạy | Trang **Lịch sử** (SQLite), ghi mọi nguồn chạy, có link xem log gốc |
| Xem kết quả phải tự mở Excel | Trang **Excel** xem trực tiếp trên web, đúng độ rộng cột như file gốc |
| Đăng nhập lại phải mở terminal | Nút **"Đăng nhập lại"** ngay trên web |
| Đặt lịch chạy phải sửa PowerShell | Tab **"Lịch chạy"** trên web, chọn giờ, bấm Lưu |

Những gì **không đổi**: cấu trúc cột Excel, cách đặt tên file PDF
(`yymmdd-soVB - tên file.pdf`), cơ chế kiểm tra trùng, cơ chế kết thúc nhanh.
Excel cũ và PDF cũ vẫn dùng bình thường; có script hỗ trợ gộp Excel cũ nếu muốn
(mục 9).

> Lưu ý an toàn: không gửi, không copy, không đưa lên GitHub file
> `playwright/.auth/state.json` – đây là file lưu phiên đăng nhập DOffice.

---

## Cài đặt nhanh

Mã nguồn và bộ phát hành nằm tại GitHub. Chọn đúng một lệnh cho máy của bạn:

### A. Máy Windows (đa số trường hợp - không cần cài Docker)

Mở PowerShell và chạy đúng một dòng:

```powershell
irm https://raw.githubusercontent.com/zerokiem/DO_Auto/v1.1.0/install.ps1 | iex
```

`install.ps1` tự tạo môi trường ảo, cài thư viện + trình duyệt Chromium (mất
vài phút, cần internet) vào `%USERPROFILE%\DO_Auto`. Có thể chọn thư mục khác
bằng biến `DOFFICE_INSTALL_DIR`. Xong nó in ra các bước tiếp theo (đăng nhập
DOffice một lần rồi chạy web). Yêu cầu duy nhất: máy đã cài sẵn **Python 3.10+**
([tải tại đây](https://www.python.org/downloads/), nhớ tick **Add python.exe
to PATH** lúc cài) - không cần cài Git, không cần Docker.

Nếu đã tải Source ZIP từ trang Release, vẫn có thể giải nén rồi chạy
`powershell -ExecutionPolicy Bypass -File .\install.ps1` như trước.

### B. Máy CÓ Docker (NAS Synology, Raspberry Pi, WSL2, Docker Desktop, Linux)

Dùng khi muốn chạy 24/7 không cần mở máy tính (xem thêm
[`README_NAS.md`](README_NAS.md) cho chi tiết Synology).

Linux headless có Docker + Compose:

```bash
curl -fsSL https://raw.githubusercontent.com/zerokiem/DO_Auto/v1.1.0/install.sh | bash
```

Synology DS423 (đổi `<user>` thành tài khoản DSM/SSH của bạn):

```bash
curl -fsSL https://raw.githubusercontent.com/zerokiem/DO_Auto/v1.1.0/install.sh | sudo env DOFFICE_INSTALL_DIR=/volume1/homes/<user>/Working/Programming/DO_Auto DOFFICE_DATA_HOST=/volume1/homes/<user>/Working/Van_ban bash
```

Script tự tải đúng release, tạo `.env`, build và bật container. Xong, mở
`http://<địa-chỉ-máy>:8877`. Dữ liệu (PDF/Excel) mặc định lưu ở `./data`; đường
dẫn truyền bằng `DOFFICE_DATA_HOST` được ghi lại vào `.env`.

### Bước bắt buộc 1 lần: đăng nhập DOffice (`state.json`)

Dù cài kiểu nào, DOffice cần 1 phiên đăng nhập đã lưu (`playwright/.auth/state.json`).

- **Máy có màn hình** (đa số trường hợp - Windows theo Cách A ở trên): vào tab
  **Cài đặt** trên web dashboard, bấm **"Đăng nhập lại"** - cửa sổ Chromium mở
  ngay trên máy đó để đăng nhập tay, xong bấm "Tôi đã đăng nhập xong" là lưu
  phiên tự động. Hoặc chạy `python login_save_state.py` nếu thích dòng lệnh.
- **Máy không màn hình** (NAS/Pi theo Cách B): vào tab **Cài đặt**, cùng chỗ đó
  sẽ hiện 2 ô **Tài khoản/Mật khẩu DOffice** - điền vào rồi bấm "Đăng nhập",
  server tự động đăng nhập giúp (chỉ hoạt động nếu trang đăng nhập DOffice là
  form tài khoản/mật khẩu đơn giản, không CAPTCHA/OTP - hiện đúng như vậy).
  Mật khẩu chỉ dùng đúng 1 lần lúc đó, không lưu lại ở bất kỳ đâu. Nếu cách này
  không chạy được (trang DOffice đổi giao diện...), vẫn có thể đăng nhập trên 1
  máy Windows khác rồi copy file `playwright/.auth/state.json` sang.

---

## 0. Cấu trúc bộ file

```text
DO_Auto/
│
├─ install.ps1                        # Cài đặt 1 lệnh trên Windows (tạo venv, cài thư viện + Chromium)
├─ install.sh                         # Cài đặt 1 lệnh bằng Docker trên Linux/NAS
├─ config.py                          # NGUỒN CẤU HÌNH DUY NHẤT - sửa tay hoặc qua web Cài đặt đều ghi vào đây
├─ run_doffice.py                     # Chạy chương trình bằng CLI/PowerShell
├─ run_web.py                         # Chạy bảng điều khiển web (mục 7)
├─ login_save_state.py                # Đăng nhập thủ công 1 lần, lưu phiên (bản CLI)
├─ get_telegram_chat_id.py            # Tìm Chat ID Telegram (mục 15.2)
├─ migrate_old_excel.py               # (tuỳ chọn) gộp 3 file Excel cũ vào file mới
├─ requirements.txt
├─ run_all_doffice.ps1                # Chạy tất cả tác vụ, dùng thủ công hoặc Task Scheduler
├─ create_doffice_task_examples.ps1   # Tạo lịch chạy tự động bằng PowerShell (thay thế: web /scheduler)
├─ install_web_startup.ps1            # Cài web dashboard tự chạy khi đăng nhập Windows (mục 7.8)
├─ Dockerfile, docker-compose.yml, .env.example, docker/  # Chạy qua Docker (NAS/Pi/WSL2...) - xem README_NAS.md
├─ README.md, README_NAS.md
│
├─ do_auto/                           # Package chứa logic dùng chung, thường KHÔNG cần đụng vào
│  ├─ __init__.py
│  ├─ task_types.py                   # Cấu trúc 1 "tác vụ" (TaskConfig)
│  ├─ text_utils.py                   # Dọn text, đặt tên file, parse khối chỉ đạo, ép UTF-8 console (Windows)
│  ├─ excel_log.py                    # Ghi Excel gộp nhiều sheet, độ rộng cột dùng chung cho web
│  ├─ browser_nav.py                  # Đăng nhập, chọn vai trò, điều hướng menu
│  ├─ extract.py                      # Trích dữ liệu từ 1 dòng văn bản
│  ├─ pdf_download.py                 # Tải PDF, đặt tên file
│  ├─ finish_doc.py                   # Bấm "Kết thúc nhanh" + "Lưu"
│  ├─ log_capture.py                  # "Tee" stdout ra file - MỌI lần chạy đều có log
│  ├─ history.py                      # Ghi/đọc lịch sử chạy (SQLite) - CLI, Scheduled Task, web dùng chung
│  ├─ notify.py                       # Gửi tin nhắn tổng kết qua Telegram sau mỗi phiên chạy (mục 15)
│  ├─ settings_store.py               # Đọc/ghi trực tiếp config.py cho trang Cài đặt trên web
│  ├─ login_flow.py                   # Đăng nhập DOffice tự web: tương tác (Windows) hoặc headless (NAS/Pi)
│  ├─ scheduler.py                    # Quản lý lịch chạy tự động - tự chọn Windows Task/thread nền (Linux)
│  ├─ inprocess_scheduler.py          # Lịch chạy tự động khi chạy Docker/Linux (thread nền, không dùng cron)
│  └─ runner.py                       # Vòng lặp xử lý + điều phối nhiều tác vụ + tạo log file mỗi lần chạy
│
└─ webapp/                            # Bảng điều khiển web (Flask), xem mục 7
   ├─ app.py                          # Route Flask: dashboard, chạy, lịch sử, Excel, cài đặt, lịch chạy, đăng nhập
   ├─ run_manager.py                  # Chạy tác vụ nền + phát nhật ký trực tiếp cho trình duyệt
   ├─ login_manager.py                # Điều phối luồng đăng nhập kích hoạt từ web
   ├─ templates/                      # Giao diện HTML (Jinja)
   └─ static/                        # CSS/JS
```

Bạn hầu như chỉ cần sửa **`config.py`** (hoặc dùng trang **Cài đặt** trên web –
cả 2 cùng ghi vào đúng 1 file, xem mục 7.5). Các file trong `do_auto/` là logic
dùng chung, chỉ cần sửa nếu DOffice đổi giao diện (xem mục 13).

---

## 1. 5 tác vụ mặc định (có thể thêm/bỏ qua Cài đặt)

| Khoá (`--tasks`) | Tên hiển thị | Menu DOffice | Có bấm Kết thúc? |
|---|---|---|---|
| `chu_tri` | Văn bản Chủ trì – Đã xử lý | Văn bản → Đã xử lý → tab Chủ trì | Không (văn bản đã xong) |
| `phoi_hop` | Văn bản Phối hợp – Chờ xử lý | Văn bản → Chờ xử lý → tab Phối hợp | Có |
| `dang_doan` | Công việc Đảng/Đoàn/Công đoàn – Chờ thực hiện | Công việc → Chờ thực hiện → tab Phối hợp | Có |
| `vb_duyet` | Văn bản đã duyệt – Đã phát hành | Văn bản → Đã phát hành | Không (văn bản đã phát hành) |
| `vb_de_biet` | Văn bản – Xem để biết | Văn bản → Xem để biết → tab Chưa xử lý | Có |

Sheet Excel tương ứng: **Chủ trì**, **Phối hợp**, **Đảng - Đoàn**, **Văn bản đã duyệt**
và **de_biet** (nếu bạn đã
có sheet với tên cũ `Chu_tri`/`Phoi_hop`/`Dang_doan` từ bản trước, lần chạy đầu
tiên với bản này sẽ **tự đổi tên sheet cũ** sang tên mới, giữ nguyên toàn bộ dữ
liệu, không tạo sheet trống mới).

`chu_tri` có cấu trúc chọn văn bản khác 2 tác vụ kia (đúng như bạn ghi chú):
danh sách "Đã xử lý" cho phép đọc đủ thông tin văn bản **ngay trên dòng danh
sách**, nên chương trình kiểm tra trùng **trước khi** mở văn bản (đỡ tốn thời
gian mở lại văn bản đã có trong Excel). Còn `phoi_hop`/`dang_doan` cần **mở văn
bản ra trước** thì phần "chỉ đạo" mới hiển thị đủ để trích. Sự khác biệt này đã
được đưa vào `config.py` qua cờ `check_duplicate_before_open`, không cần bạn tự
xử lý.

---

## 2. Yêu cầu trước khi chạy

- Windows 10/11.
- Python 3.11 hoặc 3.12.
- Quyền truy cập DOffice.
- Trình duyệt Chromium do Playwright cài.

Kiểm tra Python:

```powershell
python --version
```

## 3. Cài đặt môi trường lần đầu

```powershell
cd "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\DO_Auto"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
playwright install chromium
```

`requirements.txt` giờ có thêm `flask` (chỉ dùng cho bảng điều khiển web, mục 7).

## 4. Đăng nhập và lưu phiên DOffice lần đầu

```powershell
python login_save_state.py
```

Quy trình: Chromium mở ra → đăng nhập DOffice thủ công (chọn đúng tài
khoản/chức danh nếu DOffice yêu cầu) → quay lại PowerShell, nhấn `Enter` →
kiểm tra đã có file `playwright\.auth\state.json`.

Nếu sau này DOffice báo hết phiên / tự quay về màn hình login, chạy lại bước
này, **hoặc** dùng nút "Đăng nhập lại" ngay trên web dashboard (mục 7.7) –
không cần mở terminal.

---

## 5. Chỉnh `config.py` – việc quan trọng nhất khi dùng máy mới / người dùng khác

Mở `config.py`, có 2 phần:

### 5.1 Cấu hình chung (áp dụng mọi tác vụ)

```python
AUTH_STATE = Path("playwright/.auth/state.json")
DOWNLOAD_BASE_DIR = Path(r"D:\OneDrive - NPT\9. Jobs\Van_ban")
EXCEL_FILE = DOWNLOAD_BASE_DIR / "Tong_hop_DOffice.xlsx"
HISTORY_DB = DOWNLOAD_BASE_DIR / "doffice_auto_history.sqlite3"
LOGS_DIR = DOWNLOAD_BASE_DIR / "logs"
DUPLICATE_CHECK_MODE = "so_vb_ngay_vb"
STOP_WHEN_DUPLICATE_FOUND = True
SLOW_MO_MS = 250
PAUSE_BEFORE_CLOSE = False
ROLE_BUTTON_NAME_HINT = ""
```

Ý nghĩa từng dòng giống hệt bản cũ (xem mục 10 – các biến vẫn quen thuộc), chỉ
khác `DOWNLOAD_DIR`/`LOG_FILE` giờ dùng chung `DOWNLOAD_BASE_DIR`/`EXCEL_FILE`
cho cả 3 tác vụ, mỗi tác vụ tự có thư mục con và sheet riêng. `HISTORY_DB` và
`LOGS_DIR` là 2 biến mới, phục vụ trang Lịch sử trên web (mục 7.3).

### 5.2 Vai trò (role) – linh hoạt cho người dùng khác, không riêng gì 1 người

Đây là điểm khác biệt lớn nhất so với bản cũ. Trước đây vai trò
`"Phó Truyền tải điện"` bị viết cứng trong code; giờ mỗi tác vụ có 1 dòng
`role_pattern` ngay trong `config.py` – **hoặc chỉnh trực tiếp trên trang web
Cài đặt (mục 7.5), cả 2 cách đều ghi vào đúng 1 file `config.py`**:

```python
TASKS = {
    "chu_tri": TaskConfig(
        ...
        role_pattern="Phó Truyền tải điện",
        ...
    ),
    "phoi_hop": TaskConfig(
        ...
        role_pattern="Phó Truyền tải điện",
        ...
    ),
    "dang_doan": TaskConfig(
        ...
        role_pattern="Chi bộ 1",
        ...
    ),
}
```

**Người dùng khác chỉ cần**: đăng nhập DOffice, bấm vào nút góc trên có dạng
`"<Tên> Phòng ban:"`, xem trong menu xổ xuống hiện đúng chữ gì (ví dụ
`"Trưởng Truyền tải điện"`, `"Kỹ thuật"`, `"Chi bộ 2"`...), rồi gõ đúng cụm đó
vào `role_pattern` tương ứng (trong `config.py` hoặc trang web Cài đặt). Không
cần sửa bất kỳ file `.py` nào khác trong `do_auto/`.

- Nếu tài khoản chỉ có **1 vai trò duy nhất** (DOffice không hiện menu chọn),
  để `role_pattern=""` – chương trình sẽ bỏ qua bước chọn vai trò.
- Nút "Phòng ban:" được tìm theo mẫu chung (bất kỳ tên nào + "Phòng ban:"), nên
  **không cần** gõ tên người dùng vào đâu cả. Chỉ điền `ROLE_BUTTON_NAME_HINT`
  nếu trang có nhiều hơn 1 nút dạng này (trường hợp rất hiếm).

### 5.3 Các cờ khác trong từng tác vụ

```python
enable_finish=True/False              # có bấm "Kết thúc nhanh" + "Lưu" không
ask_confirm_before_finish=True/False  # có hỏi y/n trước khi bấm không (chỉ có tác dụng khi chạy CLI)
max_documents=...                     # số văn bản tối đa 1 lần chạy
enable_download_pdf=True/False        # có tải PDF không, hay chỉ ghi Excel
enabled=True/False                    # có nằm trong lựa chọn "Tất cả" không
```

Khuyến nghị khi test lần đầu trên máy mới:

```powershell
python run_doffice.py --tasks chu_tri --test
```

Cờ `--test` tự động ghi đè `max_documents=1`, `enable_finish=False`,
`ask_confirm_before_finish=True`, `PAUSE_BEFORE_CLOSE=True`, `SLOW_MO_MS=800`
cho **mọi** tác vụ trong lần chạy đó (không sửa `config.py`, không ảnh hưởng
lần chạy thật sau này). Trang web cũng có nút "Chế độ test an toàn" tương tự
(mục 7.2), nhưng vì web không có ai ngồi gõ `y/n`, phần hỏi xác nhận từng bước
sẽ luôn tắt dù bật test hay không.

---

## 5.4 Điều hướng DOffice và quản lý tác vụ

Trang **Cài đặt** mặc định dùng luồng ổn định sau:

1. Chọn **Vai trò**.
2. Mở **Đường dẫn trực tiếp đến danh sách**; trang web có danh sách gợi ý các
   route đã khảo sát trên DOffice nhưng vẫn cho nhập route mới.
3. Chọn **Tab bắt buộc** đối với dạng văn bản/công việc xử lý. Tab có thể là
   `Chủ trì`, `Phối hợp`, `Chưa xử lý`, `Đã xử lý`... tùy màn hình thực tế.

Đường dẫn tương ứng với `list_link_href` trong `config.py`. Sau khi chọn Vai
trò, Playwright `goto` thẳng route này và bỏ qua hai click Sidebar/Tiểu mục, do
đó không bị lỗi khi DOffice có hai menu trùng tên như `Xem để biết`. Cách cũ
`sidebar_item` + `list_link` vẫn nằm trong phần thu gọn để dự phòng khi chưa
biết URL. Dạng `published` có thể không có Tab.

Mục **Dạng dòng văn bản** có hai lựa chọn đã được chương trình hỗ trợ:

- `directive`: văn bản/công việc xử lý; tự nhận dạng để click nội dung, biểu tượng cờ hoặc Số VB ở đầu dòng Công việc.
- `published`: văn bản đã phát hành, click ô trong dòng bảng.

Với danh sách **Công việc → Xem để biết**, hệ thống đọc riêng từng trường từ
đúng phần tử trong hàng: Số VB, Ngày VB, Nơi phát hành và Trích yếu. Không dùng
toàn bộ nội dung ô nên không gộp bốn trường vào một cột Excel.

Nếu tài khoản có bước đặc biệt (bộ lọc, nút riêng...), có thể khai báo thủ công
`navigation_steps` trong `config.py`. Chuỗi JSON nâng cao không hiển thị trên
trang Cài đặt để tránh làm rối giao diện. Ví dụ:

```python
navigation_steps=[
    {"name": "Mở Công việc", "type": "sidebar", "value": "Công việc"},
    {"name": "Mở Đã giao việc", "type": "link_text", "value": "Đã giao việc"},
    {
        "name": "Chọn bộ lọc của tôi",
        "type": "selector",
        "selector": "button[data-testid='my-filter']",
        "action": "click",
        "timeout_ms": 10000,
        "delay_ms": 800,
    },
    {"name": "Chờ danh sách", "type": "selector", "selector": "tr.mat-row", "action": "wait"},
]
```

`type` hỗ trợ `sidebar`, `link_text`, `link_href`, `tab` và `selector`.
`selector` nhận selector Playwright (CSS, `text=...`, `xpath=...`...). Có thể
thêm `nth`, `optional`, `timeout_ms`, `delay_ms`; với `tab`, đặt
`empty_if_zero: true` (khi nhập JSON trên web; dùng `True` khi sửa `config.py`)
để coi tab có 0 bản ghi là thành công thay vì lỗi.

Trong mỗi tác vụ còn có `document_row_selector` (mặc định `tr.mat-row`) và
`document_click_selector` để thích nghi khi cấu trúc dòng/cú click của DOffice
khác nhau. Để trống selector click thì hệ thống dùng các selector dự phòng cũ.

Ở cuối trang Cài đặt có form **Tạo tác vụ mới**. Tác vụ mới tự có thư mục PDF và
sheet Excel riêng. Nút **Xóa tác vụ này** chỉ xóa cấu hình chạy: sheet Excel cũ
và toàn bộ dữ liệu lịch sử vẫn được giữ, đồng thời tiếp tục hiển thị ở trang
Excel để tra cứu.

---

## 6. Chạy bằng CLI/PowerShell

### 6.1 Menu chọn tác vụ (không cần nhớ tham số)

```powershell
.\.venv\Scripts\Activate.ps1
python run_doffice.py
```

```text
=== DOffice Auto - Chọn công việc cần xử lý ===
  1. Văn bản Chủ trì - Đã xử lý
  2. Văn bản Phối hợp - Chờ xử lý
  3. Công việc Đảng/Đoàn/Công đoàn - Chờ thực hiện (Phối hợp)
  4. Tất cả (1 + 2 + 3)
  0. Thoát

Nhập số (có thể chọn nhiều, cách nhau bằng dấu phẩy, ví dụ 1,3):
```

Gõ `4` để chạy cả 3, hoặc ví dụ `1,3` để chạy 2 trong 3 tác vụ. Chương trình
chỉ mở/đăng nhập trình duyệt **1 lần** dù chọn bao nhiêu tác vụ.

### 6.2 Chạy trực tiếp bằng tham số (không cần menu – dùng cho script/scheduler)

```powershell
python run_doffice.py --all
python run_doffice.py --tasks chu_tri,phoi_hop
python run_doffice.py --tasks dang_doan --no-pause
python run_doffice.py --list        # xem danh sách tác vụ hiện có
python run_doffice.py --all --source scheduler   # đánh dấu trong Lịch sử là chạy từ Task Scheduler
```

### 6.3 Chạy tất cả bằng `run_all_doffice.ps1`

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\run_all_doffice.ps1"
```

File này gọi `python run_doffice.py --all --no-pause --source scheduler`, có
khoá chống chạy chồng (lock file) và ghi log vào `scheduler_logs\` (log của
riêng PowerShell, **ngoài** file log mà `run_doffice.py` tự tạo trong
`LOGS_DIR` – xem mục 11). Muốn đổi danh sách tác vụ chạy tự động, sửa biến
`$RunnerArgs` đầu file, ví dụ:

```powershell
$RunnerArgs = @("--tasks", "chu_tri,phoi_hop", "--no-pause", "--source", "scheduler")
```

---

## 7. Bảng điều khiển web

```powershell
.\.venv\Scripts\Activate.ps1
python run_web.py
```

Server chạy tại `http://127.0.0.1:8877` và tự mở trình duyệt sau ~1 giây.
`python run_web.py` chạy **ở tiền cảnh** (foreground) của cửa sổ terminal đang
mở nó – đóng cửa sổ đó (hoặc `Ctrl+C`) là server tắt theo. Muốn server tự chạy
liên tục kể cả khi không mở terminal, xem mục 7.8.

> **Đây là công cụ nội bộ, dùng 1 mình (hoặc vài người tin cậy qua Tailscale)**
> – trang web KHÔNG có màn hình đăng nhập riêng. Không port-forward ra
> internet công khai. Muốn dùng từ điện thoại/laptop khác trong cùng
> Tailscale, mở `run_web.py`, đổi `HOST = "127.0.0.1"` thành `HOST = "0.0.0.0"`,
> chạy lại, rồi vào địa chỉ Tailscale của máy này, ví dụ `http://100.x.y.z:8877`.

### 7.1 Trang Bảng điều khiển (`/`)

- Mỗi tác vụ hiện là 1 thẻ: đường dẫn menu DOffice, vai trò đang dùng, số văn
  bản tối đa, và **lần chạy gần nhất** (thành công/lỗi, thời điểm, số văn bản
  ghi mới) – lấy từ trang Lịch sử.
- Tick chọn 1, vài, hoặc cả 3 tác vụ rồi bấm **"Chạy các tác vụ đã chọn"**.
- Có banner cảnh báo nếu **chưa có phiên đăng nhập** hoặc **phiên đã lưu quá 7
  ngày**, cùng nút **"Đăng nhập lại"** ngay bên dưới (mục 7.7).

### 7.2 Chạy ẩn (headless) hay hiện cửa sổ?

- **Chạy ẩn** (mặc định khi dùng web): Chromium chạy ngầm, không hiện cửa sổ.
  Phù hợp khi bạn kích hoạt từ xa (vd từ điện thoại qua Tailscale) vì không có
  ai ngồi xem cửa sổ trên máy chạy server.
- **Chạy hiện cửa sổ**: bỏ tick "Chạy ẩn" nếu bạn đang ngồi ngay máy chạy
  server và muốn quan sát/debug trực tiếp như khi chạy CLI.
- **Quan trọng**: cửa sổ Chromium (nếu chọn hiện) luôn mở trên **máy đang chạy
  `run_web.py`**, không phải máy bạn đang mở trình duyệt xem dashboard.
- **Chạy trên NAS/Docker/Pi (không có màn hình)**: ô "Chạy ẩn" bị khoá cố định
  ở trạng thái bật (server tự ép headless=True bất kể có tick hay không), vì
  máy đó không có màn hình để hiện cửa sổ - tránh lỗi/crash nếu lỡ bỏ tick.

### 7.3 Nhật ký trực tiếp + trang Lịch sử

Khi bấm chạy, khung "Nhật ký chạy trực tiếp" hiện ngay dưới bảng chọn tác vụ,
log giống hệt khi chạy CLI, cập nhật theo thời gian thực (Server-Sent Events).
Mở lại trang giữa chừng 1 lượt chạy vẫn tự nối lại đúng luồng log đang chạy.

**Mọi lần chạy đều tự tạo 1 file log riêng trong `LOGS_DIR`** – không chỉ khi
chạy qua Task Scheduler như phiên bản trước. Dù chạy tay (`python
run_doffice.py`), qua Task Scheduler, hay bấm nút trên web, `runner.py` đều
tự mở 1 file `.log` mới (tên dạng `20260713_090000_web.log`, hậu tố là nguồn
chạy) và ghi lại toàn bộ log ngay khi chạy, không cần cấu hình gì thêm.

Trang **Lịch sử** (`/history`) liệt kê 150 lần chạy gần nhất – **gộp cả chạy
tay (`cli`), Task Scheduler (`scheduler`), lẫn web (`web`)** vào 1 chỗ (cột
"Nguồn"), vì cả 3 đều ghi vào cùng 1 file `HISTORY_DB` mỗi khi `runner.py`
được gọi. Mỗi dòng có link **"Xem log"** mở đúng file log gốc của lần chạy đó
ngay trên web, không cần mở thư mục `LOGS_DIR` bằng tay.

### 7.4 Trang Excel (`/excel`)

Xem trực tiếp dữ liệu 3 sheet (mặc định 50 dòng gần nhất mỗi sheet, mới nhất ở
trên - chọn lại 20/50/100/200/300/500 dòng hoặc "Tất cả" bằng dropdown), có ô
tìm kiếm lọc theo bất kỳ cột nào, và nút **Tải file Excel** để tải nguyên file
`.xlsx` về xem đầy đủ/chỉnh sửa. Bảng có khung cuộn riêng (không cuộn theo cả
trang) với hàng tiêu đề dính ở trên. Cột "Tên file lưu" là **link bấm mở được
trực tiếp**: chạy Windows sẽ dùng đúng đường dẫn lưu; Docker/NAS sẽ dùng URL
web/Tailscale đặt trong `.env`. Độ rộng từng cột trên web **dùng đúng số liệu độ rộng cột đã đặt cho
file Excel** (`do_auto/excel_log.py`, biến `COLUMN_WIDTHS` – 1 nguồn duy nhất
cho cả Excel lẫn web, không bị lệch); bảng có thể cuộn ngang nếu màn hình hẹp
hơn tổng độ rộng các cột.

### 7.5 Trang Cài đặt (`/settings`)

Chỉnh nhanh các mục hay đổi nhất **mà không cần mở code**:

- Chung: dừng khi gặp trùng, khoá kiểm tra trùng, tốc độ thao tác trình duyệt.
- Từng tác vụ: bật/tắt, **vai trò (role_pattern)**, số văn bản tối đa, có tải
  PDF không, có bấm Kết thúc không, có hỏi xác nhận trước khi Kết thúc không.
- **Nơi lưu file**: điền thư mục gốc nơi máy chạy DOffice lưu PDF,
  Excel và log. Với Docker/NAS,
  chỉ dùng đường dẫn bên trong container (thường `/data`); muốn đổi vị trí thật
  trên NAS thì đổi `DOFFICE_DATA_HOST` trong `.env` và khởi động lại container.
- **Phiên đăng nhập DOffice**: xem mục "Bước bắt buộc 1 lần" ở đầu file này -
  máy có màn hình thì bấm nút mở Chromium thật; máy không màn hình (NAS/Pi)
  thì điền form tài khoản/mật khẩu ngay tại đây.

**Lưu sẽ ghi TRỰC TIẾP vào `config.py`** (thay thế đúng dòng cần đổi bằng biểu
thức tìm-thay-thế có mục tiêu, giữ nguyên mọi comment/định dạng khác trong
file), rồi nạp lại cấu hình ngay trong tiến trình web đang chạy – không cần
khởi động lại server. Vì `config.py` giờ là **nguồn duy nhất**: mở file bằng
tay sửa cũng thấy đúng trên web, sửa trên web cũng thấy đúng khi mở file bằng
tay hoặc chạy CLI/Task Scheduler lần sau. Các mục điều hướng menu (sidebar,
tab...) vẫn phải sửa trong `config.py` vì sai sót ở đây có thể làm hỏng hẳn
tác vụ, không đưa lên form web.

### 7.6 Trang Lịch chạy (`/scheduler`)

Đặt lịch chạy tự động (tác vụ "Tất cả") mà không cần sửa PowerShell tay: chọn
tối đa **3 thời điểm** trong ngày, bấm **Lưu**. Để trống cả 3 ô rồi Lưu sẽ xoá
lịch chạy tự động. Cơ chế lưu lịch **tự chọn theo môi trường đang chạy**
(`do_auto/scheduler.py`), trang sẽ hiện đúng banner cho biết đang dùng cách nào:

- **Windows**: xoá lịch cũ (nếu có) và tạo lại **1 Windows Scheduled Task duy
  nhất** tên `DOffice Auto Schedule`, gắn nhiều trigger lên cùng 1 task (1 task
  có thể có nhiều thời điểm kích hoạt, không cần tạo nhiều task riêng lẻ). Việc
  này gọi PowerShell (`Register-ScheduledTask`) - nếu gặp lỗi quyền truy cập
  khi lưu, thử chạy `python run_web.py` với quyền Administrator. Muốn đặt lịch
  bằng tay thay vì qua web, `create_doffice_task_examples.ps1` vẫn dùng được
  (tạo cùng 1 task tên giống hệt, nên web và PowerShell không xung đột nhau).
- **NAS/Docker/Linux**: không dùng Windows Task Scheduler (không có) cũng
  không dùng cron - lưu giờ vào file `schedule_times.txt` trong thư mục code,
  1 thread nền ngay trong tiến trình web (`do_auto/inprocess_scheduler.py`) tự
  kiểm tra mỗi 20 giây. Chạy được liên tục miễn container còn hoạt động
  (`restart: unless-stopped`), không cần dashboard đang mở trình duyệt xem.

### 7.7 Đăng nhập lại từ web

Khi phiên đăng nhập hết hạn hoặc vừa đổi mật khẩu, bấm **"Đăng nhập lại"**
trên trang Bảng điều khiển thay vì phải mở terminal:

1. Bấm "Đăng nhập lại" → 1 cửa sổ Chromium thật mở ra **trên máy đang chạy
   server** (không phải máy bạn đang xem trình duyệt, giống lưu ý ở mục 7.2).
2. Đăng nhập DOffice thủ công trong cửa sổ đó như bình thường.
3. Quay lại trang web, bấm **"Tôi đã đăng nhập xong, lưu phiên"**.
4. Hệ thống lưu `playwright/.auth/state.json` và đóng cửa sổ Chromium.

Vẫn không thể tự động hoá hoàn toàn bước đăng nhập (cần 1 người gõ tài
khoản/mật khẩu thật) – nút này chỉ giúp không phải rời trang web để mở
terminal chạy `login_save_state.py`.

### 7.8 Chạy liên tục + tự khởi động cùng Windows

`python run_web.py` bản thân nó **nhẹ khi không có tác vụ nào đang chạy** (chỉ
là 1 tiến trình Flask, không mở Chromium cho tới khi bạn bấm "Chạy" hoặc "Đăng
nhập lại") – để mở cả ngày không tốn tài nguyên đáng kể. Chromium chỉ được mở
(và tốn CPU/RAM) trong lúc thực sự đang xử lý văn bản hoặc đang đăng nhập, y
hệt như chạy CLI thủ công.

Muốn dashboard **luôn sẵn sàng** để bạn mở trình duyệt xem bất cứ lúc nào,
không cần nhớ mở terminal trước:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\install_web_startup.ps1"
```

Script này tạo 1 Scheduled Task ("DOffice Web Dashboard") chạy `pythonw.exe
run_web.py` (không hiện cửa sổ console) mỗi khi bạn đăng nhập Windows. Chạy
ngay không cần đợi đăng nhập lại:

```powershell
Start-ScheduledTask -TaskName "DOffice Web Dashboard"
```

Dừng dashboard đang chạy ngầm (không có nút Stop riêng vì không có console để
bấm):

```powershell
Get-Process pythonw | Stop-Process
```

Gỡ bỏ tự khởi động:

```powershell
Unregister-ScheduledTask -TaskName "DOffice Web Dashboard" -Confirm:$false
```

---

## 8. Tạo Scheduled Task tự động bằng PowerShell (thay thế: mục 7.6)

Cách thuận tiện hơn là dùng trang web **Lịch chạy** (mục 7.6). Nếu muốn làm
bằng tay/không mở web:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\create_doffice_task_examples.ps1"
```

Script tạo **1 task duy nhất** tên `DOffice Auto Schedule` với nhiều trigger
(mặc định 02:30, 12:45, 18:15 – sửa mảng `$Times` đầu file để đổi). Vì dùng
đúng tên task như trang web dùng, 2 cách này không xung đột – tạo bằng
PowerShell rồi vẫn xem/sửa lại được trên web, và ngược lại.

Kiểm tra / chạy thử / xoá task:

```powershell
Get-ScheduledTask -TaskName "DOffice Auto Schedule"
Start-ScheduledTask -TaskName "DOffice Auto Schedule"
Get-ScheduledTaskInfo -TaskName "DOffice Auto Schedule"
Unregister-ScheduledTask -TaskName "DOffice Auto Schedule" -Confirm:$false
```

---

## 9. Gộp dữ liệu Excel cũ (tuỳ chọn, chỉ chạy 1 lần)

Nếu bạn đã dùng 3 script cũ và muốn giữ lại lịch sử trong 3 file Excel cũ
(`Tong_hop_VB_chu_tri_da_XL.xlsx`, `Tong_hop_VB_phoi_hop.xlsx`,
`Tong_hop_VB_Dang_doan_phoi_hop.xlsx`) vào file gộp mới:

1. Mở `migrate_old_excel.py`, sửa `OLD_FILES` cho đúng đường dẫn 3 file cũ trên
   máy bạn.
2. Chạy:

   ```powershell
   python migrate_old_excel.py
   ```

3. Script tự thêm từng dòng cũ vào đúng sheet tương ứng (tên tiếng Việt mới)
   trong `config.EXCEL_FILE`, **tự bỏ qua** dòng đã trùng (theo cùng khoá kiểm
   tra trùng trong `config.DUPLICATE_CHECK_MODE`) nên chạy lại nhiều lần vẫn an
   toàn, không bị nhân đôi dữ liệu.

Nếu không cần giữ lịch sử cũ, bỏ qua bước này – các sheet mới sẽ tự tạo trống
trong lần chạy đầu tiên.

(Nếu bạn từng chạy 1 bản gộp cũ hơn với sheet tên `Chu_tri`/`Phoi_hop`/
`Dang_doan`, không cần làm gì thêm – bản này tự đổi tên sheet cũ sang tên tiếng
Việt mới trong lần chạy đầu tiên, xem mục 1.)

---

## 10. Ý nghĩa các biến cấu hình (tham khảo nhanh)

| Biến | Ý nghĩa |
|---|---|
| `DOWNLOAD_BASE_DIR` | Thư mục gốc chứa PDF; mỗi tác vụ có 1 thư mục con (`download_subdir`) |
| `DOWNLOAD_BASE_DIR_OVERRIDE` | Thư mục lưu nhập trên trang Cài đặt; áp dụng cho PDF, Excel, lịch sử và log. Để trống sẽ dùng biến môi trường/mặc định. |
| `EXCEL_FILE` | 1 file Excel duy nhất, mỗi tác vụ 1 sheet (`sheet_name`) |
| `HISTORY_DB` | File SQLite lưu lịch sử chạy, dùng cho trang Lịch sử trên web |
| `LOGS_DIR` | Thư mục chứa file log của TỪNG LẦN CHẠY (mọi nguồn: cli/scheduler/web) |
| `MAX_DOCUMENTS` (`max_documents` trong từng tác vụ) | Số văn bản tối đa 1 lần chạy |
| `ENABLE_DOWNLOAD_PDF` (`enable_download_pdf`) | Có tải PDF hay chỉ ghi Excel |
| `ENABLE_FINISH_DOCUMENT` (`enable_finish`) | Có bấm Kết thúc nhanh + Lưu không |
| `ASK_CONFIRM_BEFORE_FINISH` (`ask_confirm_before_finish`) | Hỏi y/n trước khi Kết thúc/Lưu (chỉ có tác dụng khi chạy CLI) |
| `STOP_WHEN_DUPLICATE_FOUND` | Dừng khi gặp văn bản đã có trong Excel |
| `DUPLICATE_CHECK_MODE` | `"so_vb_ngay_vb"` (khuyến nghị) / `"so_vb"` / `"so_vb_time"` |
| `REFRESH_LIST_EVERY` | 0 = không tự load lại danh sách; >0 = load lại sau N văn bản |
| `SLOW_MO_MS` | Tốc độ thao tác trình duyệt, tăng để dễ quan sát khi test |
| `PAUSE_BEFORE_CLOSE` | Dừng chờ Enter trước khi đóng browser (luôn `False` khi chạy qua web) |
| `role_pattern` | Tên chức danh cần chọn trên DOffice cho từng tác vụ (mục 5.2) |
| `DOWNLOAD_BASE_DIR` (env `DOFFICE_DATA_DIR`) | Chạy Windows: đường dẫn trong `config.py`. Chạy Docker/NAS: đọc từ biến môi trường `DOFFICE_DATA_DIR` (`docker-compose.yml`), thường là `/data` bên trong container – xem `README_NAS.md`. |
| `DISPLAY_BASE_URL` (env `DOFFICE_DISPLAY_URL`) | Chỉ dùng khi chạy Docker/NAS: URL public kết thúc bằng `/vb` để hyperlink PDF mở qua web/Tailscale. Chạy Windows không cần đặt, link `file://` sẽ trỏ thẳng tới đúng nơi lưu. |

Gợi ý cấu hình theo chế độ (giống bản cũ):

**Test cực an toàn** (dùng `--test` hoặc nút test trên web là đủ, không cần tự set):
`max_documents=1`, `enable_finish=False`, `ask_confirm_before_finish=True`,
`PAUSE_BEFORE_CLOSE=True`, `SLOW_MO_MS=800`.

**Chạy tự động bằng Scheduled Task**:
`ask_confirm_before_finish=False`, `PAUSE_BEFORE_CLOSE=False`, `SLOW_MO_MS=0`
(có thể sửa trực tiếp trong `config.py`, trang web Cài đặt, hoặc truyền `--no-pause`).

---

## 11. Kết quả đầu ra

- PDF: `DOWNLOAD_BASE_DIR\<download_subdir>\yymmdd-soVB - tên file.pdf`
- Excel: 1 file `EXCEL_FILE`, 3 sheet `Chủ trì` / `Phối hợp` / `Đảng - Đoàn`, cột
  giống hệt bản cũ (STT, Số VB, Ngày VB, Nơi phát hành, Trích yếu, Người chỉ
  đạo, Thời gian chỉ đạo *(ẩn)*, Nội dung chỉ đạo, Chủ trì, Phối hợp, Thư mục
  lưu, Tên file lưu *(có hyperlink)*, Thời gian lưu). Xem nhanh trên web ở
  trang Excel (mục 7.4).
- Log: `LOGS_DIR\<yyyymmdd_HHMMSS>_<nguồn>.log` – **1 file cho MỌI lần chạy**,
  bất kể chạy tay/Task Scheduler/web. Xem trực tiếp qua link "Xem log" ở trang
  Lịch sử, hoặc mở file text bình thường.
- Lịch sử chạy: `HISTORY_DB` (SQLite), xem ở trang Lịch sử trên web hoặc mở
  bằng bất kỳ công cụ SQLite nào.
- Ảnh debug khi lỗi: `debug_<tiền tố tác vụ>_<loại lỗi>.png` trong thư mục chạy
  script, ví dụ `debug_vb_phoi_hop_download_failed.png`.

---

## 12. Lỗi thường gặp

| Triệu chứng | Cách xử lý |
|---|---|
| `Không thấy file session: playwright/.auth/state.json` | Chạy `python login_save_state.py`, hoặc bấm "Đăng nhập lại" trên web (mục 7.7) |
| Script tự quay về màn hình đăng nhập | Phiên hết hạn, đăng nhập lại (CLI hoặc web) |
| Không ghi được Excel | Đóng file `Tong_hop_DOffice.xlsx` đang mở rồi chạy lại |
| Không click được nút Tải xuống / Kết thúc / Lưu | Xem ảnh debug tương ứng; kiểm tra DOffice có đổi giao diện, có đổi tên nút không |
| Chọn nhầm/không chọn được vai trò | Kiểm tra lại `role_pattern` trong `config.py` (hoặc trang Cài đặt) có đúng chữ hiển thị trong menu "...Phòng ban:" không (mục 5.2) |
| Scheduled Task chạy nhưng không thấy kết quả | `Get-ScheduledTaskInfo`, kiểm tra `$ProjectDir`/`$PythonExe` trong `run_all_doffice.ps1`, và `PAUSE_BEFORE_CLOSE` phải là `False`. Xem file log mới nhất trong `LOGS_DIR` hoặc trang Lịch sử trên web để biết chi tiết lỗi |
| Trang web báo "Chưa có phiên đăng nhập" dù đã chạy `login_save_state.py` | Kiểm tra `AUTH_STATE` trong `config.py` có trỏ đúng đường dẫn không; đường dẫn tương đối tính từ thư mục đang chạy `run_web.py`/`run_doffice.py` |
| Bấm "Chạy" trên web nhưng không thấy cửa sổ Chromium đâu | Đang bật "Chạy ẩn (headless)" – tắt tick đó nếu muốn thấy cửa sổ, hoặc xem log trực tiếp trong khung Nhật ký thay vì nhìn cửa sổ |
| 2 người cùng bấm "Chạy" trên web cùng lúc | Web chỉ cho 1 lượt chạy hoạt động tại 1 thời điểm (giống lock file của `run_all_doffice.ps1`); lượt sau sẽ báo lỗi "Đang có 1 lượt chạy khác" |
| Lưu Cài đặt trên web báo lỗi "Không lưu được" | Đóng mọi chương trình đang mở `config.py` (vd VSCode có thể khoá file trên 1 số hệ thống); kiểm tra thông báo lỗi chi tiết hiển thị trên trang |
| Trang "Lịch chạy" báo lỗi khi Lưu | Trên Windows: tính năng gọi PowerShell (`Register-ScheduledTask`) - thử chạy `python run_web.py` với quyền Administrator. Trên NAS/Docker: kiểm tra thư mục code có ghi được file `schedule_times.txt` không |
| Đăng nhập headless (NAS/Pi) báo "không tìm thấy ô tài khoản/mật khẩu" | DOffice có thể đã đổi giao diện trang đăng nhập - xem `do_auto/login_flow.py` (các selector đang dùng) hoặc quay lại đăng nhập trên máy Windows rồi copy `state.json` |

---

## 13. Khi DOffice đổi giao diện – sửa ở đâu

DOffice là web app động, có thể đổi HTML/selector sau khi cập nhật. Nhờ đã gộp
logic dùng chung, giờ chỉ cần sửa **1 chỗ** thay vì sửa lặp lại ở 3 file (và
cả CLI lẫn web dashboard tự động dùng bản sửa mới, vì cả 2 cùng gọi chung
`do_auto/`):

| Thay đổi trên DOffice | Sửa ở file |
|---|---|
| Đổi cấu trúc HTML dòng văn bản (So VB, Ngày VB, Trích yếu, khối chỉ đạo...) | `do_auto/extract.py` |
| Đổi tên/vị trí nút Tải xuống, PDF viewer | `do_auto/pdf_download.py` |
| Đổi tên/vị trí nút Kết thúc nhanh / Lưu | `do_auto/finish_doc.py` |
| Đổi cách vào sidebar / link / tab | `do_auto/browser_nav.py` (hàm `click_sidebar_item`, `click_tab`) |
| Đổi cách chọn chức danh | `do_auto/browser_nav.py` (hàm `choose_role_if_needed`) — thường chỉ cần sửa `role_pattern` trong `config.py`/trang Cài đặt, không cần sửa code |

Khi sửa, nên test với `--test` (giới hạn 1 văn bản, hỏi xác nhận từng bước)
trước khi chạy thật.

---

## 15. Thông báo qua Telegram (tuỳ chọn)

Sau mỗi phiên chạy (dù chạy tay, Task Scheduler, hay web), hệ thống có thể tự
gửi 1 tin nhắn Telegram tổng kết: tác vụ nào chạy, bao nhiêu văn bản mới, số
VB + trích yếu của từng văn bản mới (tối đa 10 dòng/tác vụ), và báo lỗi nếu có
tác vụ nào thất bại.

**Vì sao chọn Telegram**: miễn phí hoàn toàn, không giới hạn, API chính thức
đơn giản (1 request HTTP là gửi được tin nhắn), không cần xác minh doanh
nghiệp như Zalo Official Account hay WhatsApp Business API (2 dịch vụ này khó
dùng cho 1 người/nội bộ, phải qua duyệt của Meta/Zalo). Các thư viện
WhatsApp/Zalo "không chính thức" (đăng nhập bằng quét QR như Baileys) vi phạm
điều khoản dịch vụ và dễ bị khoá tài khoản, nên không dùng ở đây.

### 15.1 Tạo bot Telegram (1 lần duy nhất)

1. Mở Telegram, tìm và nhắn tin cho **@BotFather**.
2. Gửi lệnh `/newbot`, đặt tên bot, đặt username (phải kết thúc bằng `bot`, vd
   `doffice_binh_bot`).
3. BotFather trả về 1 chuỗi **token** dạng
   `<BOT_TOKEN>` – copy lại và giữ bí mật.
4. Tìm đúng bot vừa tạo (theo username đã đặt), nhắn 1 tin bất kỳ (vd "hi") để
   bot "biết" cuộc trò chuyện với bạn tồn tại. Muốn nhận tin trong 1 nhóm thay
   vì tin nhắn riêng, add bot vào nhóm rồi nhắn 1 tin trong nhóm đó.

### 15.2 Lấy Chat ID

```powershell
python get_telegram_chat_id.py
```

Script hỏi Bot Token (hoặc tự đọc từ `config.py` nếu đã điền sẵn), gọi
Telegram để tìm các cuộc trò chuyện gần đây, in ra Chat ID tương ứng (số dương
= nhắn riêng, số âm = trong nhóm).

### 15.3 Bật thông báo

Khuyến nghị điền vào `.env` (đặc biệt khi chạy Docker/NAS hoặc lưu code trên
GitHub):

```dotenv
DOFFICE_TELEGRAM_ENABLED=true
DOFFICE_TELEGRAM_BOT_TOKEN=<BOT_TOKEN>
DOFFICE_TELEGRAM_CHAT_ID=111222333
DOFFICE_TELEGRAM_ONLY_IF_NEW=false
```

**Hoặc** vào trang web **Cài đặt** &rsaquo; mục "Thông báo qua Telegram", điền
Bot Token + Chat ID, bấm **"Gửi thử"** để kiểm tra ngay trước khi lưu (không
cần đợi chạy thật 1 tác vụ mới biết cấu hình đúng hay sai), rồi bấm "Lưu cài
đặt" — ghi vào cấu hình máy đang chạy như mọi mục khác (mục 7.5). Không commit
token hoặc file `.env` lên GitHub.

### 15.4 Hành vi

- Mặc định (`TELEGRAM_NOTIFY_ONLY_IF_NEW = False`): gửi tin nhắn sau **mọi**
  phiên chạy, kể cả khi không có văn bản mới nào — để biết chắc automation vẫn
  đang chạy bình thường, không âm thầm hỏng (vd DOffice đổi giao diện làm
  script không mở được danh sách vẫn sẽ báo lỗi qua Telegram).
- Đặt `TELEGRAM_NOTIFY_ONLY_IF_NEW = True` nếu chạy tự động nhiều lần/ngày và
  chỉ muốn được thông báo khi thực sự có văn bản mới, đỡ nhiễu.
- Chạy ở `--test`/chế độ test trên web vẫn gửi thông báo bình thường, có gắn
  thêm nhãn `[CHẾ ĐỘ TEST]` đầu tin nhắn để phân biệt với phiên chạy thật.
- Gửi thất bại (sai token, mất mạng...) chỉ in cảnh báo ra log, **không bao
  giờ** làm hỏng hay dừng phiên chạy đang xử lý văn bản.
- Tin nhắn dài quá giới hạn Telegram (4096 ký tự) sẽ tự cắt bớt, kèm ghi chú
  xem đầy đủ trong Excel/trang Lịch sử.

---

## 16. Giới hạn hiện tại của bảng điều khiển web

Ghi chú thẳng thắn để bạn không bất ngờ khi dùng:

- **Không có xác thực đăng nhập riêng.** Ai vào được địa chỉ web (localhost
  hoặc qua Tailscale) đều bấm "Chạy"/"Đăng nhập lại"/sửa "Cài đặt"/"Lịch chạy"
  được. Chỉ nên bật `HOST = "0.0.0.0"` khi chắc chắn mọi thiết bị trong
  Tailscale của bạn đều đáng tin.
- **Chỉ 1 lượt chạy tự động hoá tại 1 thời điểm** (khoá bằng in-memory lock
  trong tiến trình Flask), không có hàng đợi (queue) nhiều lượt chạy nối tiếp
  nhau. Luồng "Đăng nhập lại" dùng khoá riêng, không tranh chấp với luồng chạy
  tác vụ (2 việc có thể diễn ra cùng lúc nếu thật sự cần, dù hiếm khi cần thiết).
- **`Flask dev server`** (`app.run(...)`) dùng để đơn giản hoá triển khai. Với
  quy mô dùng nội bộ 1-vài người, server này đủ ổn định; nếu cần chắc chắn hơn,
  có thể thay bằng `waitress-serve webapp.app:app --port=8877`.
- **Nhật ký trực tiếp dùng chung `sys.stdout`** cho cả tiến trình trong lúc
  đang chạy 1 tác vụ – vì chỉ cho phép 1 lượt chạy cùng lúc nên không có
  nguy cơ trộn log giữa 2 lượt chạy khác nhau, nhưng nếu bạn tự thêm code
  `print()` ở nơi khác trong lúc đang chạy, dòng đó cũng sẽ lọt vào khung
  Nhật ký và vào file log của lần chạy đó.
- **Đăng nhập headless (NAS/Pi) chỉ chạy được nếu DOffice vẫn dùng form
  tài khoản/mật khẩu đơn giản** (không CAPTCHA/OTP/SSO redirect) - nếu DOffice
  đổi sang xác thực nhiều lớp, cách này sẽ báo lỗi rõ ràng và bạn cần quay lại
  đăng nhập trên máy Windows rồi copy `state.json` (xem mục "Bước bắt buộc 1
  lần" ở đầu file).
