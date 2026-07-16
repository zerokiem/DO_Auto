# DOffice Auto – Tự động tải và xử lý văn bản DOffice bằng Playwright

## 1. Giới thiệu

Script Python tự động thao tác trên hệ thống EVN Digital Office (DOffice):

* Đăng nhập và lưu session
* Vào mục Văn bản → Chờ xử lý → Phối hợp
* Tự động mở văn bản đầu tiên
* Tải file PDF
* Kết thúc văn bản
* Ghi log tổng hợp vào file Excel

Công nghệ sử dụng:

* Python
* Playwright
* openpyxl

---

# 2. Cấu trúc thư mục

Ví dụ:

```text
DO_Auto/
│
├─ .venv/
├─ playwright/
│  └─ .auth/
│     └─ state.json
│
├─ login_save_state.py
├─ doffice_auto_v3_excel.py
├─ requirements.txt
└─ README.md
```

---

# 3. Cài đặt môi trường

## 3.1 Cài Python

Khuyến nghị:

* Python 3.11 hoặc 3.12

Kiểm tra:

```powershell
python --version
python -m pip install --upgrade pip
```

---

## 3.2 Tạo virtual environment

Trong thư mục project:

```powershell
python -m venv .venv
```

Kích hoạt:

```powershell
.venv\Scripts\activate
```

---

## 3.3 Cài thư viện

```powershell
pip install playwright openpyxl
playwright install chromium
```

---

# 4. Đăng nhập lần đầu (BẮT BUỘC)

## Chạy file:

```powershell
python login_save_state.py
```

Script sẽ:

1. Mở trình duyệt Chromium
2. Vào trang DOffice
3. Người dùng đăng nhập thủ công
4. Sau khi đăng nhập thành công:

   * quay lại PowerShell
   * nhấn Enter

Session đăng nhập sẽ được lưu tại:

```text
playwright/.auth/state.json
```

Sau này script sẽ dùng lại session này để không cần đăng nhập nữa.

---

# 5. Chạy script tự động xử lý văn bản

## Chạy:

```powershell
python doffice_auto_v3_excel.py
```

Script sẽ:

1. Vào Văn bản → Chờ xử lý → Phối hợp
2. Chọn văn bản đầu tiên
3. Tải PDF
4. Kết thúc văn bản
5. Ghi log Excel
6. Lặp tiếp

---

# 6. Các tham số cấu hình quan trọng

Trong file:

```text
doffice_auto_v3_excel.py
```

## 6.1 Thư mục lưu văn bản

```python
DOWNLOAD_DIR = Path(r"D:\OneDrive - NPT\9. Jobs\Van_ban")
```

Đổi theo nhu cầu.

Ví dụ:

```python
DOWNLOAD_DIR = Path(r"D:\VB_DOffice")
```

---

## 6.2 File Excel tổng hợp

```python
LOG_FILE = DOWNLOAD_DIR / "Tong_hop_van_ban_da_xu_ly.xlsx"
```

Có thể đổi tên file nếu muốn.

---

## 6.3 Số lượng văn bản xử lý mỗi lần chạy

```python
MAX_DOCUMENTS = 4
```

Ví dụ:

```python
MAX_DOCUMENTS = 20
```

---

## 6.4 Có tự động kết thúc văn bản hay không

```python
ENABLE_FINISH_DOCUMENT = True
```

### True

* Có bấm “Kết thúc văn bản”

### False

* Chỉ tải file
* Không kết thúc

---

## 6.5 Có yêu cầu xác nhận trước khi kết thúc hay không

```python
ASK_CONFIRM_BEFORE_FINISH = False
```

### True

Mỗi văn bản sẽ hỏi:

```text
Chuẩn bị bấm KẾT THÚC văn bản này. Đồng ý? y/n:
```

### False

Tự động xử lý hoàn toàn.

---

## 6.6 Tốc độ thao tác trình duyệt

```python
SLOW_MO_MS = 250
```

* Đơn vị: milliseconds

Ví dụ:

```python
SLOW_MO_MS = 1000
```

→ thao tác chậm hơn để dễ quan sát.

---

# 7. Chỉnh chức danh xử lý văn bản

Trong hàm:

```python
def choose_role_if_needed(page):
```

Có đoạn:

```python
role = page.get_by_role(
    "menuitem",
    name=re.compile(r"Phó Truyền tải điện", re.I)
)
```

Đổi theo chức danh tương ứng.

Ví dụ:

```python
name=re.compile(r"Trưởng Truyền tải điện", re.I)
```

hoặc:

```python
name=re.compile(r"Kỹ thuật", re.I)
```

---

# 8. Dữ liệu được ghi log Excel

Mỗi văn bản sẽ ghi:

| STT | Nội dung           |
| --- | ------------------ |
| 1   | Số VB              |
| 2   | Ngày VB            |
| 3   | Nơi phát hành      |
| 4   | Trích yếu          |
| 5   | Người chỉ đạo      |
| 6   | Thời gian chỉ đạo  |
| 7   | Nội dung chỉ đạo   |
| 8   | Chủ trì thực hiện  |
| 9   | Phối hợp thực hiện |
| 10  | Thư mục lưu        |
| 11  | Tên file lưu       |
| 12  | Thời gian lưu      |

File Excel:

```text
Tong_hop_van_ban_da_xu_ly.xlsx
```

---

# 9. Lưu ý quan trọng

## 9.1 Không mở file Excel khi script đang chạy

Nếu file Excel đang mở bằng Microsoft Excel:

* script có thể không ghi được dữ liệu
* báo lỗi PermissionError

---

## 9.2 Session đăng nhập có thể hết hạn

Nếu script tự nhiên bị đá ra login:

* xóa file:

```text
playwright/.auth/state.json
```

* chạy lại:

```powershell
python login_save_state.py
```

---

## 9.3 Không commit session lên GitHub

Không upload:

```text
playwright/.auth/state.json
```

vì file này chứa session đăng nhập.

---

# 10. Debug khi lỗi

Nếu script lỗi:

* sẽ tự động chụp ảnh màn hình debug:

Ví dụ:

```text
debug_download_failed.png
debug_finish_button_failed.png
```

Dùng ảnh này để kiểm tra trạng thái giao diện lúc lỗi.

---

# 11. Requirements.txt (tham khảo)

Có thể tạo file:

```text
requirements.txt
```

Nội dung:

```text
playwright
openpyxl
```

Cài nhanh:

```powershell
pip install -r requirements.txt
```

---

# 12. Workflow khuyến nghị

## Giai đoạn test

```python
MAX_DOCUMENTS = 1
ENABLE_FINISH_DOCUMENT = False
ASK_CONFIRM_BEFORE_FINISH = True
```

## Giai đoạn vận hành thật

```python
MAX_DOCUMENTS = 20
ENABLE_FINISH_DOCUMENT = True
ASK_CONFIRM_BEFORE_FINISH = False
```

---

# 13. Chức năng hiện tại

✅ Tự động đăng nhập bằng session
✅ Tự động chọn đúng chức danh
✅ Tự động vào tab Phối hợp
✅ Tự động tải PDF
✅ Tự động kết thúc văn bản
✅ Tự động ghi log Excel
✅ Tự động append log cũ
✅ Xử lý nhiều văn bản liên tục

---

# 14. Hướng nâng cấp tương lai

* Đặt tên file PDF theo:

  * số văn bản
  * trích yếu
  * ngày văn bản

* Tự động phân thư mục:

  * theo năm
  * theo đơn vị phát hành
  * theo loại văn bản

* Chạy scheduler tự động mỗi ngày

* Tích hợp OCR/PDF parsing

* Dashboard thống kê tiến độ xử lý văn bản
