# DOffice Auto - Context cho Claude Code

## Tiến độ hiện tại
Đã gộp 3 script Playwright riêng (Chủ trì/Phối hợp/Đảng-Đoàn xử lý văn bản DOffice) thành 1 dự án thống nhất `DO_Auto/`, gồm:
- CLI (`run_doffice.py`) + PowerShell scheduler
- Web dashboard Flask (`run_web.py`) tại `http://127.0.0.1:8877`: chạy tác vụ, xem log trực tiếp (SSE), lịch sử (SQLite), xem Excel, cài đặt, đăng nhập lại, đặt lịch Task Scheduler
- Thông báo Telegram sau mỗi phiên chạy
- Vừa fix xong 1 bug nghiêm trọng: `config.py` cũ của user thiếu field mới → web sập 500. Đã sửa bằng `getattr()` phòng thủ + auto-append field thiếu vào `config.py` khi lưu Settings.

Toàn bộ đã test qua bản sao cô lập (không phải test thật trên Windows/DOffice thật — sandbox lúc phát triển không có Windows/PowerShell/mạng nội bộ DOffice).

## Kiến trúc / Setup
```
DO_Auto/
├─ config.py              # NGUỒN CẤU HÌNH DUY NHẤT (sửa tay hoặc qua web đều ghi vào đây)
├─ run_doffice.py          # CLI entry
├─ run_web.py              # Web entry (Flask, port 8877)
├─ login_save_state.py     # Đăng nhập thủ công lưu session
├─ get_telegram_chat_id.py
├─ migrate_old_excel.py
├─ run_all_doffice.ps1 / create_doffice_task_examples.ps1 / install_web_startup.ps1
├─ do_auto/                # package logic dùng chung
│  ├─ task_types.py        # TaskConfig, TaskResult (dataclass)
│  ├─ text_utils.py, excel_log.py, browser_nav.py, extract.py
│  ├─ pdf_download.py, finish_doc.py, runner.py (orchestrator chính)
│  ├─ history.py (SQLite), log_capture.py, notify.py (Telegram)
│  ├─ settings_store.py    # đọc/ghi config.py bằng regex thay thế, giữ format
│  ├─ login_flow.py, scheduler.py (Windows Task Scheduler qua PowerShell subprocess)
└─ webapp/
   ├─ app.py, run_manager.py, login_manager.py
   ├─ templates/*.html, static/app.js, style.css
```
- Deps: `playwright`, `openpyxl`, `flask` (requirements.txt)
- Windows 10/11, Python 3.11/3.12, venv `.venv`
- Excel: 1 file `Tong_hop_DOffice.xlsx`, 3 sheet "Chủ trì"/"Phối hợp"/"Đảng - Đoàn"

## Vấn đề / Bug hiện tại
- **Chưa test thật trên máy Windows của user** — mọi test đều mô phỏng (mock Playwright/PowerShell) trong môi trường phát triển không có Windows/DOffice/Telegram thật. Cần user tự chạy thật để phát hiện lỗi môi trường thực (selector DOffice sai, PowerShell permission, v.v.)
- Bug vừa sửa (config.py thiếu field mới → 500) — **đã fix nhưng user chưa xác nhận đã áp dụng bản vá và chạy lại thành công**.
- `/scheduler` (Windows Task Scheduler qua PowerShell subprocess) chưa được xác nhận chạy thật trên Windows.
- Web dashboard không có auth riêng — chỉ nên chạy local/Tailscale tin cậy.

## Quyết định đã chốt
- **Không dùng JSON overlay riêng** (đã bỏ `user_settings.json`) — `config.py` là single source of truth, web Settings sửa trực tiếp vào đó bằng regex thay thế có mục tiêu (giữ nguyên comment/format), không phải ghi đè cả file.
- **Mọi field đọc từ config.py phải qua `getattr()` với default** trong `settings_store.build_effective_config()` — không bao giờ được đọc trực tiếp `config.XXX` ở bất kỳ đâu khác (đây là nguyên nhân bug vừa rồi, phải giữ kỷ luật này cho mọi field mới sau này).
- Chọn **Telegram** cho thông báo (không phải Zalo/WhatsApp) vì free, API đơn giản, không cần duyệt doanh nghiệp.
- `run_selected_tasks()` trong `runner.py` là **entry point duy nhất** mọi nơi gọi vào (CLI/scheduler/web) — log file, ghi history, gửi Telegram đều xảy ra ở đây để đảm bảo nhất quán.
- Web chạy Flask dev server (không Docker/Nginx) vì Playwright cần Chromium thật trên máy, không hợp container hoá.
- 1 Windows Scheduled Task có thể gắn **nhiều trigger** (không cần nhiều task riêng) — áp dụng cho `/scheduler`.

## File / Link quan trọng
- `config.py` — sửa role/path/Telegram token ở đây (hoặc qua web `/settings`)
- `do_auto/settings_store.py` — logic đọc/ghi config.py, chỗ hay phát sinh bug loại "thiếu field"
- `do_auto/runner.py` — orchestrator chính, mọi tính năng mới nên hook vào đây
- `webapp/app.py` — toàn bộ route Flask
- `README.md` (trong project) — đã viết đầy đủ hướng dẫn setup/troubleshoot, dùng làm tài liệu tham khảo chính

## Bước tiếp theo
1. Chạy thật trên Windows với DOffice thật để xác nhận: selector Playwright còn đúng, PowerShell Scheduler tạo task thành công, Telegram gửi được.
2. Xác nhận bản vá `settings_store.py` / `app.py` / `run_doffice.py` / `run_manager.py` (4 file vá gần nhất) đã chạy ổn, dashboard load lại được.
3. Nếu cần thêm tính năng, giữ nguyên convention: field mới trong `config.py` → thêm vào `_DEFAULTS` trong `settings_store.py` → dùng `getattr()` khi đọc → không đọc `config.XXX` trực tiếp ở nơi khác.
