# DOffice Auto trên NAS Synology DS423 (Docker)

Trạng thái: **đã triển khai và chạy thật thành công** (kiểm tra ngày 16/07/2026 —
chạy thử tác vụ Chủ trì, tải PDF, ghi Excel, gửi Telegram đều OK). Tài liệu này
vừa là **nhật ký triển khai** vừa là **hướng dẫn vận hành/khôi phục** nếu cần
làm lại từ đầu trên NAS khác.

Container `doffice` chạy trên **DS423** (ARM Realtek RTD1619B, 2GB RAM, DSM 7.2),
lưu code + văn bản ngay trên NAS. Truy cập qua LAN `http://192.168.1.254:8877`
hoặc Tailscale `http://100.100.1.254:8877` (cả 2 đã test OK).

| Thành phần | Trên NAS |
|---|---|
| Code | `/volume1/homes/binhnx/Working/Programming/DO_Auto` → mount vào container tại `/app` |
| Văn bản (PDF/Excel/log/DB) | `/volume1/homes/binhnx/Working/Van_ban` → mount vào `/data` |
| Web dashboard | container `doffice`, cổng `8877`, `restart: unless-stopped` |
| Lịch chạy tự động | **thread Python nền** ngay trong `run_web.py` (KHÔNG dùng cron) |
| Đăng nhập DOffice | làm trên Windows, copy `state.json` lên NAS (NAS không có màn hình) |

---

## Những gì đã sửa trong code

- `config.py`: `DOWNLOAD_BASE_DIR` ưu tiên đọc biến môi trường `DOFFICE_DATA_DIR`
  (trong container = `/data`). Không đặt biến này (vd chạy trên Windows) thì vẫn
  dùng đường dẫn cũ `D:\OneDrive - NPT\9. Jobs\Van_ban` → **1 file dùng được cả
  2 nơi**.
- `do_auto/runner.py`: Chromium thêm `--no-sandbox --disable-dev-shm-usage` khi
  chạy trên Linux (bắt buộc trong Docker).
- `do_auto/login_flow.py`: nút "Đăng nhập lại" báo lỗi rõ ràng khi chạy trên NAS
  (hướng dẫn đăng nhập trên Windows) thay vì crash khó hiểu.
- `do_auto/notify.py`: **bỏ cắt 70 ký tự trích yếu** trong tin Telegram (trả về
  nguyên văn); tin quá dài giới hạn Telegram (4096 ký tự) tự tách thành nhiều
  tin nối tiếp thay vì cắt mất nội dung.
- **Lịch chạy tự động trên Linux đổi hẳn sang thread Python nội bộ** (không
  dùng cron/Task Scheduler nào cả) — xem `do_auto/inprocess_scheduler.py`:
  - Container trên NAS chạy 24/7 (`restart: unless-stopped`) nên không cần 1
    tiến trình cron tách biệt.
  - `run_web.py` khi chạy trên Linux tự khởi động 1 thread nền, mỗi 20 giây so
    khớp giờ hiện tại (giờ NAS, `TZ=Asia/Ho_Chi_Minh`) với danh sách giờ đã lưu
    trong `schedule_times.txt` (nằm trong thư mục code, giữ nguyên qua các lần
    container restart). Khớp giờ → gọi thẳng `RunManager.start()` (cùng cơ chế
    khoá "chỉ 1 lượt chạy tại 1 thời điểm" như nút "Chạy" trên web).
  - `do_auto/scheduler.py` giữ nguyên API cũ (`get_current_times` /
    `apply_schedule` / `remove_schedule`) nên `webapp/app.py` (route
    `/scheduler`) và giao diện web **không đổi gì** — chỉ đổi triển khai bên
    dưới theo hệ điều hành (Windows vẫn dùng Task Scheduler qua PowerShell như
    cũ).
  - `webapp/run_manager.py`: `RunManager.start()`/`_run()` nhận thêm tham số
    `trigger_source` để lượt chạy tự động ghi đúng nhãn "scheduler" trong Lịch
    sử (trước đó bị hard-code "web").
- Thêm: `Dockerfile`, `docker-compose.yml`, `docker/entrypoint.sh`,
  `.dockerignore`, `do_auto/inprocess_scheduler.py`.

### Vì sao không dùng cron/apt-get trong image

Lúc build lần đầu, `apt-get install cron tzdata` liên tục timeout khi kết nối
`ports.ubuntu.com` — nhưng CHỈ lỗi khi build qua BuildKit (`docker build`), còn
`docker run ... apt-get update` chạy tay lại OK. Đây là do BuildKit trên NAS này
dùng network namespace riêng bị lỗi định tuyến/DNS, không phải lỗi mạng NAS nói
chung. Thử ép `RUN --network=host` cũng bị chặn (`network.host is not allowed`
— entitlement này bị khoá theo mặc định, không nên mở trên NAS sản xuất).

Giải pháp cuối cùng: **bỏ hẳn `apt-get`**, không cần luôn:
- `tzdata` đã có sẵn dữ liệu zoneinfo trong base image Playwright — chỉ cần
  `ln -snf /usr/share/zoneinfo/$TZ /etc/localtime`, không cần cài gói.
- `cron` được thay bằng thread Python nội bộ (xem phần trên) — không cần cài
  gì cả.

### Vì sao phải `pip install playwright==1.55.0` dù dùng image Playwright

Image `mcr.microsoft.com/playwright/python:v1.55.0-noble` có sẵn **browser
Chromium** (tại `/ms-playwright`, biến môi trường `PLAYWRIGHT_BROWSERS_PATH`)
nhưng **không có sẵn gói pip `playwright`** — phải tự `pip install`. Ghim đúng
version `1.55.0` khớp với tag image để dùng thẳng browser có sẵn, không phải
tải lại (không gọi `playwright install`).

---

## Cách triển khai lại từ đầu (nếu cần làm trên NAS khác / cài lại)

### Bước 0 — Cấp quyền Docker không cần mật khẩu cho user SSH

Docker socket trên Synology chỉ `root` truy cập được. SSH vào NAS (interactive,
không qua script tự động) rồi chạy:

```bash
ssh binhnx@192.168.1.254
echo 'binhnx ALL=(ALL) NOPASSWD: /usr/local/bin/docker, /usr/local/bin/docker-compose, /usr/local/bin/dockerd' | sudo tee /etc/sudoers.d/binhnx-docker
sudo chmod 440 /etc/sudoers.d/binhnx-docker
sudo -n docker ps   # phải chạy KHÔNG hỏi mật khẩu
```

### Bước 1 — Đăng nhập DOffice trên Windows (tạo phiên đăng nhập)

```powershell
.\.venv\Scripts\Activate.ps1
python login_save_state.py
```

Đăng nhập DOffice trong cửa sổ Chromium hiện ra → quay lại PowerShell nhấn
`Enter`. Tạo ra file **`playwright\.auth\state.json`** — mang lên NAS ở bước
sau. **Không** đưa file này lên GitHub/chia sẻ.

### Bước 2 — Copy code + dữ liệu lên NAS

```bash
ssh binhnx@192.168.1.254 "mkdir -p /volume1/homes/binhnx/Working/Programming/DO_Auto /volume1/homes/binhnx/Working/Van_ban"
```

Cách nhanh nhất (từ Git Bash/WSL trên Windows, tại thư mục dự án): đóng gói
bằng `tar` rồi bơm qua `ssh` (tránh scp từng file nhỏ, nhanh hơn nhiều):

```bash
tar --exclude='.venv' --exclude='__pycache__' --exclude='.git' -czf - \
  config.py run_web.py run_doffice.py requirements.txt Dockerfile docker-compose.yml \
  do_auto webapp docker playwright \
  | ssh binhnx@192.168.1.254 "cd /volume1/homes/binhnx/Working/Programming/DO_Auto && tar -xzf -"
```

> **Lưu ý OneDrive Files On-Demand**: nếu thư mục văn bản cũ
> (`D:\OneDrive - NPT\9. Jobs\Van_ban`) có file "online-only" (chưa tải về máy),
> `tar` sẽ phải tải từng file từ OneDrive trước khi nén — có thể RẤT CHẬM với
> nhiều file (thực tế: >1000 file, ~1.1GB mất hàng chục phút). Cân nhắc chạy
> việc này lúc rảnh/ban đêm, hoặc chuột phải chọn "Always keep on this device"
> cho thư mục đó trước để tải sẵn về máy rồi mới nén.

### Bước 3 — Build và chạy container

```bash
ssh binhnx@192.168.1.254
cd /volume1/homes/binhnx/Working/Programming/DO_Auto
sudo -n /usr/local/bin/docker-compose build
sudo -n /usr/local/bin/docker-compose up -d
```

> NAS này có `docker-compose` (binary kiểu v1, nhưng thực chất chạy engine
> Compose v2 bên dưới — `docker-compose version` báo v2.20.x) tại
> `/usr/local/bin/docker-compose`, KHÔNG có plugin `docker compose` (v2 CLI
> lệnh liền, 2 từ). Luôn dùng `docker-compose` (gạch nối) trên NAS này.

Kiểm tra:

```bash
sudo -n /usr/local/bin/docker ps --filter name=doffice
sudo -n /usr/local/bin/docker logs doffice --tail 50
```

### Bước 4 — Mở web dashboard

- LAN: **http://192.168.1.254:8877**
- Tailscale: **http://100.100.1.254:8877**

Cổng 8877 không trùng dashboard DSM (5000/5001). **Không** port-forward 8877 ra
internet — web không có màn hình đăng nhập riêng.

---

## Đăng nhập lại khi hết phiên

NAS không có màn hình nên nút "Đăng nhập lại" trên web sẽ báo lỗi có hướng dẫn
(không dùng được ở đây). Khi DOffice báo hết phiên:

1. Trên **Windows**: `python login_save_state.py` → đăng nhập lại → có
   `state.json` mới.
2. Copy đè lên NAS:

   ```powershell
   scp "playwright/.auth/state.json" binhnx@192.168.1.254:/volume1/homes/binhnx/Working/Programming/DO_Auto/playwright/.auth/
   ```

   Không cần restart container — lần chạy sau tự dùng file mới.

---

## Đặt lịch chạy tự động

Vào tab **Lịch chạy** trên web → chọn tối đa 2 thời điểm (HH:MM, giờ GMT+7) →
**Lưu**. Container tự kiểm tra mỗi 20 giây và tự kích hoạt đúng giờ (thread
Python nội bộ, xem giải thích ở trên) — không cần cron, không cần thao tác gì
thêm trên NAS. Để trống cả 2 ô rồi Lưu = xoá lịch.

⚠️ **Lưu ý quan trọng**: đặt lịch nghĩa là container sẽ **tự động đăng nhập
DOffice thật và ghi dữ liệu thật** (kể cả bấm "Kết thúc" cho tác vụ Phối hợp/
Đảng-Đoàn nếu `enable_finish=True`) đúng giờ đã chọn — không phải hành động vô
hại. Chỉ đặt lịch khi đã chắc chắn cấu hình đúng.

Danh sách giờ lưu tại `schedule_times.txt` trong thư mục code — giữ nguyên qua
các lần container restart (vì thư mục code được mount, không nằm trong image).

---

## Vận hành thường ngày

```bash
cd /volume1/homes/binhnx/Working/Programming/DO_Auto

sudo -n /usr/local/bin/docker-compose restart          # khởi động lại (sau khi sửa code .py)
sudo -n /usr/local/bin/docker-compose down              # dừng hẳn container
sudo -n /usr/local/bin/docker-compose up -d              # chạy lại
sudo -n /usr/local/bin/docker-compose logs --tail=100    # 100 dòng log gần nhất
sudo -n /usr/local/bin/docker stats doffice               # theo dõi RAM/CPU (chỉ 2GB)
```

- **Sửa cấu hình** (role, số văn bản, Telegram...): dùng tab **Cài đặt** trên
  web — ghi thẳng vào `config.py`, có hiệu lực ngay, không cần restart.
- **Sửa code Python** (`do_auto/`, `webapp/`): sửa file trên NAS rồi
  `docker-compose restart` để nạp lại.
- **Đổi Dockerfile** (thêm thư viện hệ thống...): `docker-compose build` rồi
  `up -d`.

---

## Lưu ý RAM 2GB (DS423)

- Nên để `max_documents` vừa phải, tránh mở quá nhiều văn bản 1 lượt.
- Chạy **1 lượt tại 1 thời điểm** (khoá sẵn trong `RunManager`, cả nút Chạy lẫn
  lịch tự động đều dùng chung khoá này).
- Theo dõi bằng `sudo -n /usr/local/bin/docker stats doffice`. Nếu Chromium hay
  crash, tăng `shm_size` trong `docker-compose.yml` (vd `512mb`) rồi
  `docker-compose up -d`.

---

## Xử lý sự cố

| Triệu chứng | Cách xử lý |
|---|---|
| Web không mở được | `docker ps` xem container còn chạy không; `docker logs doffice` xem lỗi |
| Container cứ restart liên tục (`Restarting`) | `docker logs doffice` — thường là traceback Python lúc import, kiểm tra module thiếu (vd quên `pip install playwright`) |
| Banner "Chưa có phiên đăng nhập" | `state.json` chưa đúng chỗ: phải ở `DO_Auto/playwright/.auth/state.json` trong thư mục code đã mount |
| `apt-get`/`docker build` timeout tới ports.ubuntu.com | Đừng dùng apt-get trong Dockerfile trên NAS này (xem giải thích ở trên); nếu bắt buộc cần gói hệ thống mới, thử `RUN --network=host` (cần bật entitlement, có thể bị chặn) hoặc tải binary tĩnh qua HTTPS thay vì apt |
| `network.host is not allowed` khi build | BuildKit trên NAS không cho phép entitlement `network.host` theo mặc định — không cố mở, tìm cách khác (xem trên) |
| Tab "Lịch chạy" đặt xong không chạy đúng giờ | Kiểm tra `docker exec doffice date` ra đúng giờ GMT+7; kiểm tra `schedule_times.txt` trong thư mục code có đúng giờ đã lưu; xem `docker logs doffice` quanh giờ đó có dòng `[scheduler]` không |
| Sửa `config.py` trên web báo không lưu được | Thư mục code trên NAS phải cho container ghi (thuộc user `binhnx`, quyền ghi) |
| `sudo` hỏi mật khẩu dù đã cấu hình NOPASSWD | Phải dùng đúng đường dẫn tuyệt đối trong lệnh (`/usr/local/bin/docker`, không phải `docker` trần) — sudoers rule ghim theo path tuyệt đối |

---

## Ghi chú kỹ thuật

- Code **không** được COPY vào image mà **mount lúc chạy** (`docker-compose.yml`)
  → sửa code trên NAS không cần build lại image, chỉ cần restart container.
- `docker/entrypoint.sh` chỉ đơn giản `exec python run_web.py` — không còn khởi
  động cron/daemon nào khác.
- `run_doffice.py` / `run_web.py` (chạy tay hoặc do lịch tự động kích hoạt) đều
  gọi chung `runner.run_selected_tasks()` nên log, lịch sử (SQLite), Telegram
  nhất quán như trên Windows.
- Các file `.ps1` (Windows Task Scheduler) vẫn nằm trong repo để dùng khi chạy
  trên Windows, nhưng **không liên quan** đến NAS — trên NAS lịch chạy hoàn
  toàn do `do_auto/inprocess_scheduler.py` đảm nhiệm.
- File `do_auto/config.py` (khác với `config.py` ở gốc project) là bản dư thừa
  không được import ở đâu — có thể xoá an toàn, không ảnh hưởng NAS hay Windows.
