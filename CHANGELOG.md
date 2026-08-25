# Changelog

## v1.1.4 - 2026-08-25

### Added

- Bo sung bo script CMD cho Windows 10: cai moi truong, dang nhap, chay web,
  cai/go Scheduled Task va kiem tra cong 8877.
- Them huong dan cai lai sach trong `README_WINDOWS_CMD.md`.

### Fixed

- Sua ghi log khi web chay bang `pythonw.exe` khong co `sys.stdout`, tranh loi
  `'NoneType' object has no attribute 'write'`.
- Neo duong dan `AUTH_STATE` theo thu muc project de CMD, Scheduled Task va
  PowerShell dung chung mot phien dang nhap.

## v1.1.0 - 2026-08-16

### Added

- Cho phép thêm, sửa và xóa tác vụ DOffice trên trang Cài đặt; tác vụ mới tự có
  sheet Excel và thư mục PDF riêng.
- Hỗ trợ điều hướng bằng đường dẫn trực tiếp sau khi chọn Vai trò, kèm Tab Văn
  bản bắt buộc cho tác vụ dạng xử lý.
- Bổ sung tác vụ mặc định `vb_de_biet` cho Văn bản - Xem để biết.
- Thêm kiểm thử hồi quy cho tác vụ động, sheet Excel và dòng Công việc/Xem để
  biết.
- Thêm bộ cài một dòng cho Windows (`install.ps1`) và Linux/NAS Synology
  (`install.sh`), tự tải đúng release `v1.1.0`.

### Changed

- Gom các biểu mẫu tác vụ vào mục **Các tác vụ** và thu gọn từng biểu mẫu để
  trang Cài đặt dễ sử dụng hơn.
- Đổi nhãn thành **Tab Văn bản** và **Xóa tác vụ này**.
- Xóa tác vụ chỉ xóa cấu hình chạy (và vì vậy không còn trong lựa chọn Tất cả);
  sheet Excel cùng dữ liệu lịch sử luôn được giữ lại.
- Bỏ giao diện JSON selector/lưu trữ nâng cao khỏi trang Cài đặt. Cấu hình cũ
  vẫn được tương thích và có thể chỉnh thủ công trong `config.py` khi cần.
- Đưa Bot Token/Chat ID Telegram mặc định sang biến môi trường để không đóng
  gói thông tin xác thực vào source/release public.

### Fixed

- Sửa click dòng **Công việc → Xem để biết**: click trực tiếp vào Số VB để mở
  đúng tài liệu/PDF thay vì chỉ focus hàng.
- Sửa trích xuất metadata cho hàng **Xem để biết**. Số VB, Ngày VB, Nơi phát
  hành và Trích yếu nay được lấy từ bốn selector riêng, hỗ trợ cả hai biến thể
  DOM hiện gặp của DOffice và có fallback regex theo thứ tự Số VB → Ngày VB →
  Nơi phát hành → Ngày nhận → Trích yếu; không còn gộp vào một ô Excel.
- Tự nhận diện dòng **Xem để biết** ngay cả khi tác vụ đang chọn nhầm dạng
  "Văn bản đã phát hành", tránh fallback gộp toàn bộ hàng vào Số VB/Trích yếu.
- Bổ sung fallback cho nút **Kết thúc** trong `mat-menu-panel` bằng cả button
  và span, không phụ thuộc ID panel động của Angular.
- Sửa luồng kết thúc của **Văn bản → Xem để biết**: mở menu **ba chấm** sau
  nút Thông tin trước khi chọn menuitem **Kết thúc**, ưu tiên locator `span`
  đã kiểm chứng để tránh timeout không cần thiết, rồi bấm **Lưu** trong
  `dialog-ktnhanh-vbd`.
- Sửa kiểm tra popup **Kết thúc nhanh** khi Angular trả về nhiều text trùng
  nhau, không còn cảnh báo giả trước khi bấm Lưu.
- Sửa việc ghi `True`/`False` vào `config.py` từ trang Cài đặt, tránh lỗi
  `name 'true' is not defined`.
