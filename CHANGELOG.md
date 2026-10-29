# Changelog

## v1.1.3 — 2026-10-29

- Tự đọc giới hạn N ngày từ lỗi HTTP 400 của GDT, chia lại kỳ bị từ chối và
  áp dụng giới hạn mới cho các kỳ tiếp theo; hỗ trợ cả hóa đơn thường và máy
  tính tiền khi giới hạn thay đổi thành 5, 3, 1 hoặc số ngày khác.
- Không thử lại CAPTCHA khi lỗi do khoảng ngày; giữ các dòng đã nhận và tránh
  nhân đôi nếu phải chia lại kỳ giữa các trang. Kỳ một ngày lỗi vẫn được báo,
  không lặp vô hạn. Không ghi body lỗi HTTP thô ra log hoặc workbook.
- Thêm kiểm thử giới hạn thay đổi, ranh giới tháng, dữ liệu phân trang nhận
  một phần và tương thích đọc lỗi HTTP trên PowerShell 5.1/7.

## v1.1.2 — 2026-10-28

- Sửa truy vấn hóa đơn máy tính tiền: chia kỳ tối đa 10 ngày, lấy CAPTCHA mới
  cho từng trang và thử lại có giới hạn khi CAPTCHA bị HTTP 400.
- Giới hạn request danh sách ở 50 hóa đơn để tránh GDT HTTP 500 với `size=100`.
- Ghi lỗi khi API lặp con trỏ hoặc hết con trỏ trước khi nhận đủ `total`, giữ
  dữ liệu đã nhận để đối chiếu thay vì báo danh sách đã hoàn tất.

## v1.1.1 — 2026-10-15

Cập nhật asset ngày 16/10/2026:

- Khôi phục định dạng mẫu VBA: Consolas, màu riêng theo sheet, kiểu ô, độ rộng
  cột, chiều cao hàng, cố định khung, bộ lọc và lề in; giữ nguyên 64/34/32 cột.
- Kiểm thử hồi quy định dạng và ghi tiếp; tài nguyên mẫu được nhúng trong EXE.
- Rút gọn README; chuyển hướng dẫn chi tiết sang `docs/USAGE.md`.


- Chỉ một file EXE: nhấn đúp chạy, không cần tự giải nén hoặc cài đặt.
- Ghi tiếp vào một file Excel: tick **Ghi tiếp vào file đã chọn** (mặc định bật)
  hoặc `APPEND_OUTPUT=true` để nối hóa đơn mới vào file đã có. Hóa đơn trùng khóa
  (mẫu số, ký hiệu, số, ngày lập, MST người bán/mua) được thay bằng dữ liệu mới,
  nên resume, chạy lại cùng kỳ, tách riêng mua vào/bán ra hoặc đổi kỳ đều dùng
  chung một file mà không nhân đôi dữ liệu. Bảng `LinkTraCuu` và dòng lỗi của hóa
  đơn đã xử lý lại cũng được cập nhật theo.
- Tên file Excel mặc định của giao diện là `HoaDonDienTu.xlsx` (bỏ thời điểm chạy)
  để hợp với luồng ghi tiếp một file duy nhất.
- Icon riêng cho app (16–256 px) dùng trong Explorer, title bar và taskbar; cửa sổ
  WPF nạp icon từ payload khi khởi động.
- Build hỗ trợ ký Authenticode sẵn (`-SignCertificateThumbprint`, SHA-256 kèm mốc
  thời gian) khi đã có chứng chỉ; bản phát hành hiện tại chưa ký nên Windows có
  thể hiện cảnh báo SmartScreen — xem mục SmartScreen trong `README.md`.
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
