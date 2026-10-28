# Phân tích luồng GDT ngày 28/10/2026

Chỉ ghi lại hành vi API; không đưa HAR, tài khoản, token, cookie, CAPTCHA thật
hay dữ liệu hóa đơn vào repository.

## Bằng chứng trong HAR

- `POST /api/security-taxpayer/authenticate` trả HTTP 200. Body dùng
  `username`, `password`, `ckey`, `cvalue`; không cần đổi endpoint đăng nhập.
- `GET /api/sco-query/invoices/sold` với khoảng ngày dài trả HTTP 400:
  “Khoảng thời gian tìm kiếm không được lớn hơn 10 ngày.”
- Truy vấn SCO thành công gửi `ckey` và `cvalue` trong query string. Trang
  tiếp theo gửi cả `state` lẫn bộ CAPTCHA mới. Không dùng lại CAPTCHA đăng nhập.
- Danh sách trả `datas`, `total`, `state`, `time`. `state` là con trỏ opaque,
  cần URL-encode nguyên giá trị; không tự tạo hoặc tăng số trang thay thế.
- `query/invoices/purchase` trong HAR không gửi CAPTCHA truy vấn. Không áp
  yêu cầu CAPTCHA của SCO cho hóa đơn thông thường.
- `query/invoices/detail` dùng khóa `nbmst`, `khhdon`, `shdon`, `khmshdon`,
  khớp endpoint phục hồi chi tiết đang có. HAR không có request xuất XML,
  nên không suy luận endpoint xuất XML mới từ file này.

## Kiểm tra trực tiếp sau phân tích

Cùng tài khoản, endpoint và khoảng ngày: `size=50` và `size=15` trả HTTP 200,
`size=100` trả HTTP 500. Vì vậy giới hạn request ở 50, vẫn chấp nhận cấu hình
cũ tới 100 để không làm hỏng file cấu hình đã lưu.

## Thay đổi và kiểm thử

`Get-GdtInvoiceDateRanges` giữ ranh giới tháng và chia thêm SCO thành tối đa
10 ngày tính cả hai đầu, không trùng hoặc bỏ ngày. `Invoke-GdtInvoicePage`
lấy CAPTCHA từng trang, thử lại tối đa ba lần khi gặp HTTP 400.

Con trỏ lặp hoặc hết con trỏ trước `total` được ghi vào báo cáo lỗi danh sách;
giữ dữ liệu đã nhận để người dùng đối chiếu và tải lại phần thiếu.

`tests/Test-ScoQuery.ps1` dùng toàn bộ dữ liệu giả để kiểm tra chia kỳ, phân
trang, encode CAPTCHA/con trỏ, giới hạn trang, retry và phát hiện danh sách
thiếu. Chạy qua `tests/Run-Tests.ps1` trên PowerShell 5.1 và 7.
