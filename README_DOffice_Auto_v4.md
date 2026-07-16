# DOffice Auto – README hướng dẫn cài đặt, cấu hình, chạy thủ công và chạy tự động

Tài liệu này dùng cho bộ script tự động tải PDF, tổng hợp Excel và xử lý văn bản trên hệ thống DOffice bằng Python + Playwright.

Bộ file dự kiến gửi cho người sử dụng gồm:

```text
DO_Auto/
│
├─ DO_phoi_hop.py
├─ DO_Dang_Doan_phoi_hop_v3.py
├─ DO_chu_tri_da_XL_v3.py
├─ run_all_doffice_v4_ascii.ps1
├─ create_doffice_task_examples_v4_ascii.ps1
├─ login_save_state.py                 # nếu có, dùng để lưu phiên đăng nhập lần đầu
├─ requirements.txt                    # nếu có
└─ README.md
```

> Lưu ý an toàn: không gửi, không copy, không đưa lên GitHub file `playwright/.auth/state.json`, vì đây là file lưu phiên đăng nhập DOffice.

---

## 1. Mục đích của từng file

| File | Chức năng chính | Khu vực xử lý |
|---|---|---|
| `DO_phoi_hop.py` | Vào DOffice, mở danh sách văn bản chuyên môn ở tab Phối hợp, tải PDF, có thể kết thúc văn bản, ghi log Excel | `Văn bản` → `Chờ xử lý` → `Phối hợp` |
| `DO_Dang_Doan_phoi_hop_v3.py` | Vào phần Công việc của vai trò Chi bộ/Công đoàn, mở tab Phối hợp, tải PDF, có thể kết thúc công việc, ghi log Excel | `Công việc` → `Chờ thực hiện` → `Phối hợp` |
| `DO_chu_tri_da_XL_v3.py` | Tổng hợp lại các văn bản đã xử lý ở tab Chủ trì, tải PDF và ghi log Excel. Script này thường không bấm kết thúc vì văn bản đã xử lý rồi | `Văn bản` → `Đã xử lý` → `Chủ trì` |
| `run_all_doffice_v4_ascii.ps1` | Script PowerShell chạy tuần tự nhiều script Python, thường dùng để chạy thủ công hoặc làm action cho Windows Task Scheduler | PowerShell |
| `create_doffice_task_examples_v4_ascii.ps1` | Script PowerShell mẫu để tạo Scheduled Task tự động chạy `run_all_doffice_v4_ascii.ps1` theo lịch | Windows Task Scheduler |

---

## 2. Yêu cầu trước khi chạy

Máy tính cần có:

- Windows 10/11.
- Python 3.11 hoặc 3.12.
- Quyền truy cập DOffice.
- Trình duyệt Chromium do Playwright cài.
- Các thư viện Python:
  - `playwright`
  - `openpyxl`

Kiểm tra Python:

```powershell
python --version
```

---

## 3. Cài đặt môi trường lần đầu

Mở PowerShell tại thư mục project, ví dụ:

```powershell
cd "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
```

Tạo môi trường ảo:

```powershell
python -m venv .venv
```

Kích hoạt môi trường ảo:

```powershell
.\.venv\Scripts\Activate.ps1
```

Cài thư viện:

```powershell
python -m pip install --upgrade pip
pip install playwright openpyxl
playwright install chromium
```

Nếu có file `requirements.txt` thì có thể dùng:

```powershell
pip install -r requirements.txt
playwright install chromium
```

---

## 4. Đăng nhập và lưu phiên DOffice lần đầu

Các script dùng file session:

```text
playwright/.auth/state.json
```

Nếu chưa có file này, cần chạy script đăng nhập thủ công, ví dụ:

```powershell
python login_save_state.py
```

Quy trình:

1. Chromium mở ra.
2. Đăng nhập DOffice thủ công.
3. Chọn đúng tài khoản/chức danh nếu DOffice yêu cầu.
4. Quay lại PowerShell và nhấn `Enter` theo hướng dẫn của script.
5. Kiểm tra đã có file:

```text
playwright\.auth\state.json
```

Nếu sau này DOffice báo hết phiên, đăng nhập không vào, hoặc tự quay về màn hình login, làm lại bước này.

---

## 5. Cấu hình quan trọng trong từng script Python

Mỗi script có một khối cấu hình nằm gần đầu file, dạng:

```python
DOFFICE_URL = "https://doffice.npt.com.vn/"
AUTH_STATE = Path("playwright/.auth/state.json")
DOWNLOAD_DIR = Path(r"D:\OneDrive - NPT\9. Jobs\Van_ban\...")
LOG_FILE = DOWNLOAD_DIR / "..."
MAX_DOCUMENTS = ...
ENABLE_DOWNLOAD_PDF = True
STOP_WHEN_DUPLICATE_FOUND = True
DUPLICATE_CHECK_MODE = "so_vb_ngay_vb"
REFRESH_LIST_EVERY = 0
SLOW_MO_MS = 250
PAUSE_BEFORE_CLOSE = True
```

### 5.1 `DOWNLOAD_DIR`

Thư mục lưu PDF và file Excel tổng hợp.

Ví dụ:

```python
DOWNLOAD_DIR = Path(r"D:\OneDrive - NPT\9. Jobs\Van_ban\VB_Phoi_hop")
```

Có thể đổi sang thư mục khác:

```python
DOWNLOAD_DIR = Path(r"D:\DOffice_Output\VB_Phoi_hop")
```

Script sẽ tự tạo thư mục nếu chưa có.

### 5.2 `LOG_FILE`

File Excel tổng hợp.

Ví dụ:

```python
LOG_FILE = DOWNLOAD_DIR / "Tong_hop_VB_phoi_hop.xlsx"
```

Không nên mở file Excel này khi script đang chạy, vì Excel có thể khóa file làm Python không ghi được.

### 5.3 `MAX_DOCUMENTS`

Số văn bản tối đa xử lý trong một lần chạy.

Khi test nên để nhỏ:

```python
MAX_DOCUMENTS = 1
```

Khi chạy thật có thể tăng:

```python
MAX_DOCUMENTS = 10
```

hoặc:

```python
MAX_DOCUMENTS = 50
```

### 5.4 `ENABLE_DOWNLOAD_PDF`

Bật/tắt tải PDF.

```python
ENABLE_DOWNLOAD_PDF = True
```

- `True`: tải PDF về thư mục `DOWNLOAD_DIR`.
- `False`: không tải PDF, chỉ ghi metadata vào Excel nếu script cho phép đi tiếp.

### 5.5 `ENABLE_FINISH_DOCUMENT`

Có trong các script xử lý văn bản/công việc đang chờ.

```python
ENABLE_FINISH_DOCUMENT = True
```

- `True`: sau khi tải PDF, script tự bấm `Kết thúc văn bản` hoặc `Kết thúc nhanh`, rồi bấm `Lưu`.
- `False`: chỉ tải PDF và ghi Excel, không kết thúc trên DOffice.

Khuyến nghị khi test lần đầu:

```python
MAX_DOCUMENTS = 1
ENABLE_FINISH_DOCUMENT = False
ASK_CONFIRM_BEFORE_FINISH = True
```

Khi đã kiểm tra ổn mới chạy thật:

```python
ENABLE_FINISH_DOCUMENT = True
ASK_CONFIRM_BEFORE_FINISH = False
```

### 5.6 `ASK_CONFIRM_BEFORE_FINISH`

Yêu cầu xác nhận trước khi kết thúc văn bản/công việc.

```python
ASK_CONFIRM_BEFORE_FINISH = True
```

Khi bật, PowerShell sẽ hỏi `y/n` trước khi bấm kết thúc và/hoặc bấm lưu. Chế độ này an toàn khi test.

### 5.7 `STOP_WHEN_DUPLICATE_FOUND`

```python
STOP_WHEN_DUPLICATE_FOUND = True
```

Nếu gặp văn bản đã có trong Excel, script dừng lại. Cách này phù hợp khi danh sách DOffice sắp xếp văn bản mới nhất ở trên: gặp dòng đã tổng hợp rồi thì các dòng dưới thường cũng đã tổng hợp.

### 5.8 `DUPLICATE_CHECK_MODE`

Có 3 chế độ:

```python
DUPLICATE_CHECK_MODE = "so_vb_ngay_vb"
```

| Giá trị | Ý nghĩa | Khuyến nghị |
|---|---|---|
| `"so_vb_ngay_vb"` | Kiểm tra trùng theo Số VB + Ngày VB | Nên dùng mặc định |
| `"so_vb"` | Chỉ kiểm tra theo Số VB | Chỉ dùng khi chắc số VB không lặp |
| `"so_vb_time"` | Số VB + Thời gian chỉ đạo | Dùng khi cần phân biệt nhiều chỉ đạo trên cùng văn bản |

### 5.9 `REFRESH_LIST_EVERY`

```python
REFRESH_LIST_EVERY = 0
```

- `0`: không tự tải lại danh sách.
- Số lớn hơn 0: sau mỗi N văn bản thì load lại danh sách.

Nếu DOffice dùng danh sách ảo, xử lý nhiều văn bản mà bị mất dòng hoặc DOM cũ, có thể thử:

```python
REFRESH_LIST_EVERY = 10
```

### 5.10 `SLOW_MO_MS`

Tốc độ thao tác trình duyệt.

```python
SLOW_MO_MS = 250
```

Tăng lên để dễ quan sát:

```python
SLOW_MO_MS = 800
```

Giảm xuống để chạy nhanh hơn:

```python
SLOW_MO_MS = 0
```

### 5.11 `PAUSE_BEFORE_CLOSE`

```python
PAUSE_BEFORE_CLOSE = True
```

- `True`: chạy xong sẽ chờ nhấn `Enter` rồi mới đóng browser.
- `False`: chạy xong tự đóng browser.

Khi chạy bằng Task Scheduler nên để:

```python
PAUSE_BEFORE_CLOSE = False
```

Nếu để `True`, scheduled task có thể bị treo vì không có người nhấn `Enter`.

---

## 6. Chức danh / vai trò DOffice cần chỉnh

Các script có hàm dạng:

```python
def choose_role_if_needed(page) -> None:
```

Trong đó có đoạn chọn menu item theo regex.

### 6.1 Script văn bản chuyên môn

Với `DO_phoi_hop.py` và `DO_chu_tri_da_XL_v3.py`, thường dùng vai trò:

```python
role = page.get_by_role("menuitem", name=re.compile(r"Phó Truyền tải điện", re.I))
```

Nếu người dùng khác chạy, đổi cụm này theo chức danh hiển thị trên DOffice, ví dụ:

```python
role = page.get_by_role("menuitem", name=re.compile(r"Trưởng Truyền tải điện", re.I))
```

hoặc:

```python
role = page.get_by_role("menuitem", name=re.compile(r"Kỹ thuật", re.I))
```

### 6.2 Script Đảng đoàn / Chi bộ / Công đoàn

Với `DO_Dang_Doan_phoi_hop_v3.py`, script đang chọn vai trò:

```python
role = page.get_by_role("menuitem", name=re.compile(r"Chi bộ 1", re.I))
```

Nếu người khác dùng vai trò khác thì đổi `Chi bộ 1` theo đúng tên vai trò hiển thị.

---

## 7. Chạy từng script thủ công

Mở PowerShell trong thư mục project:

```powershell
cd "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
.\.venv\Scripts\Activate.ps1
```

Chạy văn bản chuyên môn phối hợp:

```powershell
python .\DO_phoi_hop.py
```

Chạy công việc Đảng đoàn/Chi bộ/Công đoàn phối hợp:

```powershell
python .\DO_Dang_Doan_phoi_hop_v3.py
```

Chạy tổng hợp văn bản Chủ trì đã xử lý:

```powershell
python .\DO_chu_tri_da_XL_v3.py
```

Nếu muốn chạy bằng Python trong `.venv` mà không cần activate:

```powershell
.\.venv\Scripts\python.exe .\DO_phoi_hop.py
```

---

## 8. Chạy tuần tự bằng `run_all_doffice_v4_ascii.ps1`

File `run_all_doffice_v4_ascii.ps1` dùng để chạy lần lượt nhiều script Python. Đây là file nên dùng khi muốn bấm một lần chạy hết, hoặc dùng làm chương trình chạy trong Windows Task Scheduler.

Cách chạy thủ công:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto\run_all_doffice_v4_ascii.ps1"
```

Trong file này thường cần chỉnh các biến đầu file, ví dụ:

```powershell
$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
$PythonExe  = "$ProjectDir\.venv\Scripts\python.exe"

$Scripts = @(
    "DO_phoi_hop.py",
    "DO_Dang_Doan_phoi_hop_v3.py",
    "DO_chu_tri_da_XL_v3.py"
)
```

Ý nghĩa:

| Biến | Ý nghĩa |
|---|---|
| `$ProjectDir` | Thư mục chứa các file `.py` và `.ps1` |
| `$PythonExe` | Đường dẫn Python trong môi trường ảo |
| `$Scripts` | Danh sách script Python sẽ chạy tuần tự |

Có thể tạm bỏ một script bằng cách comment dòng đó:

```powershell
$Scripts = @(
    "DO_phoi_hop.py",
    # "DO_Dang_Doan_phoi_hop_v3.py",
    "DO_chu_tri_da_XL_v3.py"
)
```

Nếu script PowerShell có ghi log, nên kiểm tra file log sau khi chạy để biết script nào chạy thành công hoặc lỗi.

---

## 9. Tạo Scheduled Task bằng `create_doffice_task_examples_v4_ascii.ps1`

File `create_doffice_task_examples_v4_ascii.ps1` là script mẫu để tạo tác vụ tự động trong Windows Task Scheduler.

Chạy PowerShell với quyền phù hợp, tốt nhất là **Run as Administrator**, rồi chạy:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto\create_doffice_task_examples_v4_ascii.ps1"
```

Trong file này thường cần chỉnh:

```powershell
$TaskName   = "DOffice Auto Daily"
$ProjectDir = "D:\OneDrive - NPT\1. binhnx Data\Business\Lap trinh\Python\DO_Auto"
$Runner     = "$ProjectDir\run_all_doffice_v4_ascii.ps1"
```

Ví dụ tạo task chạy mỗi ngày lúc 07:30:

```powershell
$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$Runner`"" `
    -WorkingDirectory $ProjectDir

$Trigger = New-ScheduledTaskTrigger -Daily -At 7:30AM

$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Run DOffice Auto scripts sequentially"
```

Ví dụ tạo task chạy khi đăng nhập Windows:

```powershell
$Trigger = New-ScheduledTaskTrigger -AtLogOn
```

Ví dụ tạo task chạy mỗi 2 giờ:

```powershell
$Trigger = New-ScheduledTaskTrigger -Once -At 7:30AM `
    -RepetitionInterval (New-TimeSpan -Hours 2) `
    -RepetitionDuration (New-TimeSpan -Hours 12)
```

Sau khi tạo, kiểm tra trong Task Scheduler:

```text
Task Scheduler → Task Scheduler Library → DOffice Auto Daily
```

Có thể bấm phải chuột → `Run` để chạy thử ngay.

---

## 10. Cấu hình cần lưu ý khi chạy bằng Task Scheduler

Khi chạy tự động, nên chỉnh trong các file Python:

```python
PAUSE_BEFORE_CLOSE = False
ASK_CONFIRM_BEFORE_FINISH = False
```

Nếu muốn an toàn tuyệt đối, trong giai đoạn đầu có thể để:

```python
ENABLE_FINISH_DOCUMENT = False
MAX_DOCUMENTS = 1
```

Sau khi xem log Excel và PDF tải về đúng rồi mới bật:

```python
ENABLE_FINISH_DOCUMENT = True
MAX_DOCUMENTS = 10
```

Không nên dùng scheduled task để tự động bấm kết thúc văn bản nếu chưa kiểm tra kỹ selector, chức danh và luồng xử lý.

---

## 11. Kiểm tra Scheduled Task bằng lệnh PowerShell

Xem task đã tạo:

```powershell
Get-ScheduledTask -TaskName "DOffice Auto Daily"
```

Chạy thử task:

```powershell
Start-ScheduledTask -TaskName "DOffice Auto Daily"
```

Xem trạng thái task:

```powershell
Get-ScheduledTaskInfo -TaskName "DOffice Auto Daily"
```

Xóa task nếu tạo sai:

```powershell
Unregister-ScheduledTask -TaskName "DOffice Auto Daily" -Confirm:$false
```

---

## 12. Kết quả đầu ra

Mỗi script sẽ tạo:

1. Thư mục PDF trong `DOWNLOAD_DIR`.
2. File Excel tổng hợp trong `LOG_FILE`.
3. Ảnh debug nếu lỗi, ví dụ:
   - `debug_vb_phoi_hop_download_failed.png`
   - `debug_dang_doan_phoi_hop_finish_button_failed.png`
   - `debug_chu_tri_da_xl_pdf_not_ready.png`

Các cột Excel thường gồm:

| Cột | Nội dung |
|---|---|
| STT | Số thứ tự |
| Số VB | Số/ký hiệu văn bản |
| Ngày VB | Ngày văn bản |
| Nơi phát hành | Đơn vị phát hành/gửi |
| Trích yếu | Nội dung trích yếu |
| Người chỉ đạo | Người giao/chỉ đạo |
| Thời gian chỉ đạo | Thời điểm chỉ đạo |
| Nội dung chỉ đạo | Nội dung xử lý |
| Chủ trì | Đơn vị/người chủ trì |
| Phối hợp | Đơn vị/người phối hợp |
| Thư mục lưu | Thư mục lưu PDF |
| Tên file lưu | Tên file PDF, có hyperlink |
| Thời gian lưu | Thời điểm script ghi Excel |

---

## 13. Quy tắc đặt tên file PDF

Các script đang tạo tên file dạng thân thiện, bỏ dấu tiếng Việt và thêm tiền tố:

```text
yymmdd-soVB - ten_file.pdf
```

Ví dụ:

```text
260525-2190 - Van ban quy dinh Cong tac phi_v1.pdf
```

Trong đó:

- `260525`: ngày chạy script theo dạng `yymmdd`.
- `2190`: phần số lấy từ `Số VB`.
- Nếu số văn bản ít hơn 4 chữ số thì tự thêm số 0 phía trước, ví dụ `25` → `0025`.

---

## 14. Quy trình khuyến nghị khi bàn giao cho người khác

### Bước 1 – Copy bộ file

Copy thư mục project sang máy mới, tối thiểu gồm:

```text
DO_phoi_hop.py
DO_Dang_Doan_phoi_hop_v3.py
DO_chu_tri_da_XL_v3.py
run_all_doffice_v4_ascii.ps1
create_doffice_task_examples_v4_ascii.ps1
README.md
```

Nếu có thì copy thêm:

```text
login_save_state.py
requirements.txt
```

Không copy file:

```text
playwright/.auth/state.json
```

### Bước 2 – Cài Python và thư viện

Làm theo mục 3.

### Bước 3 – Tạo session đăng nhập riêng

Chạy `login_save_state.py` trên máy người dùng đó để tạo `state.json` riêng.

### Bước 4 – Chỉnh cấu hình

Trong từng file `.py`, kiểm tra:

- `DOWNLOAD_DIR`
- `LOG_FILE`
- `MAX_DOCUMENTS`
- `ENABLE_DOWNLOAD_PDF`
- `ENABLE_FINISH_DOCUMENT`
- `ASK_CONFIRM_BEFORE_FINISH`
- chức danh trong `choose_role_if_needed()`

Trong file `.ps1`, kiểm tra:

- `$ProjectDir`
- `$PythonExe`
- danh sách `$Scripts`
- tên task và lịch chạy nếu dùng scheduled task.

### Bước 5 – Chạy test an toàn

Trong các script có kết thúc văn bản, chỉnh:

```python
MAX_DOCUMENTS = 1
ENABLE_FINISH_DOCUMENT = False
ASK_CONFIRM_BEFORE_FINISH = True
PAUSE_BEFORE_CLOSE = True
```

Chạy từng file Python trước.

### Bước 6 – Chạy thật

Khi kết quả đúng:

```python
ENABLE_FINISH_DOCUMENT = True
ASK_CONFIRM_BEFORE_FINISH = False
PAUSE_BEFORE_CLOSE = False
```

Sau đó chạy bằng:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\run_all_doffice_v4_ascii.ps1"
```

### Bước 7 – Tạo Scheduled Task nếu cần

Chỉnh và chạy:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\create_doffice_task_examples_v4_ascii.ps1"
```

---

## 15. Lỗi thường gặp và cách xử lý

### 15.1 Không thấy file `state.json`

Thông báo thường gặp:

```text
Không thấy file session: playwright/.auth/state.json
```

Cách xử lý:

```powershell
python login_save_state.py
```

### 15.2 Script bị đưa về màn hình đăng nhập

Session hết hạn. Chạy lại:

```powershell
python login_save_state.py
```

### 15.3 Không ghi được Excel

Nguyên nhân thường gặp: file Excel đang mở.

Cách xử lý:

1. Đóng file Excel tổng hợp.
2. Chạy lại script.

### 15.4 Không click được nút tải PDF

Kiểm tra:

- Văn bản có mở được viewer PDF không.
- DOffice có đổi giao diện không.
- Xem ảnh debug `debug_*_pdf_not_ready.png` hoặc `debug_*_download_failed.png`.

### 15.5 Không click được nút Kết thúc/Lưu

Kiểm tra:

- Có đúng vai trò không.
- Văn bản/công việc có quyền kết thúc không.
- Popup DOffice có bị che không.
- Xem ảnh debug `debug_*_finish_button_failed.png` hoặc `debug_*_save_finish_failed.png`.

### 15.6 Scheduled Task chạy nhưng không thấy kết quả

Kiểm tra:

```powershell
Get-ScheduledTaskInfo -TaskName "DOffice Auto Daily"
```

Và kiểm tra các điểm sau:

- `$ProjectDir` đúng chưa.
- `$PythonExe` đúng chưa.
- File `.venv\Scripts\python.exe` có tồn tại không.
- `PAUSE_BEFORE_CLOSE` đã để `False` chưa.
- Máy có đang đăng nhập Windows không.
- DOffice có yêu cầu xác thực lại không.
- Task có quyền chạy PowerShell không.

---

## 16. Gợi ý cấu hình theo từng chế độ

### Chế độ test cực an toàn

```python
MAX_DOCUMENTS = 1
ENABLE_DOWNLOAD_PDF = True
ENABLE_FINISH_DOCUMENT = False
ASK_CONFIRM_BEFORE_FINISH = True
PAUSE_BEFORE_CLOSE = True
SLOW_MO_MS = 800
```

### Chế độ chạy thủ công có kiểm soát

```python
MAX_DOCUMENTS = 5
ENABLE_DOWNLOAD_PDF = True
ENABLE_FINISH_DOCUMENT = True
ASK_CONFIRM_BEFORE_FINISH = True
PAUSE_BEFORE_CLOSE = True
SLOW_MO_MS = 250
```

### Chế độ chạy tự động bằng Scheduled Task

```python
MAX_DOCUMENTS = 10
ENABLE_DOWNLOAD_PDF = True
ENABLE_FINISH_DOCUMENT = True
ASK_CONFIRM_BEFORE_FINISH = False
PAUSE_BEFORE_CLOSE = False
SLOW_MO_MS = 0
```

---

## 17. Ghi chú bảo trì

DOffice là web app động, có thể đổi HTML/selector sau khi cập nhật giao diện. Nếu script đang chạy ổn rồi tự nhiên lỗi, ưu tiên kiểm tra:

1. Ảnh debug do script chụp.
2. Vai trò/chức danh đã chọn đúng chưa.
3. Menu/tab có đổi tên không.
4. Nút `Tải xuống`, `Kết thúc nhanh`, `Lưu` có đổi text hoặc vị trí không.
5. File Excel có đang mở không.
6. Session `state.json` có hết hạn không.

Khi cần sửa code, nên sửa từng script một và test với:

```python
MAX_DOCUMENTS = 1
ENABLE_FINISH_DOCUMENT = False
```

Sau khi chắc chắn mới tăng số lượng và bật kết thúc tự động.
