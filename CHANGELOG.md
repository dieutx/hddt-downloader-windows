# Changelog

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
