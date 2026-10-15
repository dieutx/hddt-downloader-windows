# Changelog

## v1.1.1 — 2026-10-15

- Chỉ một file EXE: nhấn đúp chạy, không cần tự giải nén hoặc cài đặt.
- Giao diện compact: header nhỏ, form hai cột, tiến độ và nhật ký luôn nhìn thấy;
  nút con mắt hiện/ẩn mật khẩu, giữ giá trị chỉnh sửa trong bộ nhớ.
- Nhật ký 4–6 dòng gần nhất, màu INFO/SUCCESS/WARNING/ERROR và cửa sổ xem toàn bộ.
  Kiểm thử layout ở 4 kích thước desktop, DPI 100%/125% và 4 trạng thái tải.
- Kiểm tra GitHub Releases chạy nền, xem nội dung thay đổi và tải EXE mới trực tiếp;
  kiểm tra dung lượng, PE header/checksum, giữ hỗ trợ ZIP portable cũ.
- Dữ liệu mặc định trong Documents/HDDT-Downloader; cấu hình tải/nhật ký trong
  LocalAppData/HDDT-Downloader, dùng tiếp khi di chuyển hoặc cập nhật EXE.
- Giữ dữ liệu tổng hợp/chi tiết và tiếp tục xử lý khi có lỗi từng hóa đơn;
  phục hồi chi tiết API cho XML HTTP 500/504, dừng an toàn khi chờ retry.
- Build không kèm credential, cấu hình cá nhân hay dữ liệu hóa đơn. Không đổi
  endpoint đăng nhập, tên sheet/cột hoặc thứ tự cột của workbook.

## v1.0.0

### App Windows

- Giao diện WPF cho Windows 10/11, chạy bằng `HDDT-Downloader.exe` hoặc
  `run-gui.cmd`; dùng Windows PowerShell 5.1 và .NET Framework tích hợp.
- Tải từ GDT hoặc chuyển XML có sẵn sang Excel, hiển thị nhật ký/tiến độ,
  dừng an toàn và giữ phần dữ liệu đã xử lý.
- Cài đặt nâng cao: kết nối XML ban đầu/tối đa, giãn cách request, phục hồi
  sau 429, retry, timeout, page size và tùy chọn nhật ký. Có thể ghi nhớ tham
  số trong `hddt-settings.json`; không lưu tài khoản hay mật khẩu.
- Gói portable không kèm `.env`, cấu hình cá nhân hoặc dữ liệu hóa đơn.

### Dữ liệu hóa đơn và Excel

- Khi xuất XML gặp HTTP 500, lấy dữ liệu bằng `/api/query/invoices/detail`
  của hóa đơn hiện tại, dùng token phiên đang chạy. Bổ sung tổng hợp và các
  dòng hàng hóa vào sheet chi tiết; báo cáo ghi rõ kết quả phục hồi.
- Trường API thiếu giữ trống hoặc giữ dữ liệu danh sách; không tạo XML giả.
- Đọc JSON từ byte UTF-8 để giữ tiếng Việt trên PowerShell 5.1; dùng Calibri
  trong workbook.
- Chỉ tạo hyperlink cho URL HTTP/HTTPS hợp lệ; URL mẫu trong `LinkTraCuu`
  giữ dưới dạng chữ, tránh relationship lỗi trong sheet 8.

### Kiểm thử

- Bộ kiểm thử lõi và WPF đạt trên Windows 11 / Windows PowerShell 5.1.
- Gói portable đã qua kiểm tra khởi tạo giao diện và cài đặt nâng cao.
- Chưa kiểm thử trên Windows 10, PowerShell 7 hoặc GDT thật trong lần phát hành.
