# HDDT Downloader

Tải XML hóa đơn điện tử **mua vào / bán ra** từ GDT và xuất Excel theo mẫu
[TaiHoaDonDienTu](https://github.com/dieutx/TaiHoaDonDienTu).
Có ứng dụng Windows một file EXE và dòng lệnh trên Windows, macOS, Linux.

**[Tải v1.1.2 cho Windows](https://github.com/dieutx/hddt-downloader-windows/releases/download/v1.1.2/HDDT-Downloader-v1.1.2-windows.exe)**
· [Release & checksum](https://github.com/dieutx/hddt-downloader-windows/releases/tag/v1.1.2)
· [Video hướng dẫn](https://www.youtube.com/watch?v=nLWvosy-IM0)

![Giao diện HDDT Downloader](docs/images/hddt-desktop-v1.1.1.png)

## Bắt đầu trên Windows

1. Tải và mở **HDDT-Downloader-v1.1.2-windows.exe**; không cần cài đặt hay cài Excel.
2. Nhập tài khoản GDT, chọn khoảng ngày và loại hóa đơn.
3. Chọn nơi lưu rồi bấm **Bắt đầu**. Bấm **Dừng an toàn** để xuất phần đã xử lý.
4. Có XML sẵn: dùng tab **XML có sẵn → Excel**, không cần đăng nhập hay Internet.

Ứng dụng dùng Windows PowerShell 5.1 và .NET Framework có sẵn trên Windows 10/11.
EXE hiện chưa ký số; Windows có thể hiện SmartScreen. Với file tải từ release
ở trên, mở **Properties → Unblock** nếu cần.

## Kết quả Excel

Truy vấn hóa đơn máy tính tiền (`sco-query`) tự chia khoảng ngày thành các kỳ
tối đa 10 ngày và lấy CAPTCHA mới cho từng trang. Hóa đơn thông thường vẫn
chia theo tháng. API trả lại con trỏ phân trang cũ hoặc kết thúc trước `total`
sẽ được ghi thành lỗi danh sách; dữ liệu đã nhận vẫn được giữ lại.
Số hóa đơn mỗi trang gửi lên GDT được giới hạn ở 50, kể cả khi cấu hình cũ
đặt `PAGE_SIZE=100`, vì máy chủ trả HTTP 500 với trang 100 hóa đơn.

Giữ font **Consolas**, màu riêng cho từng sheet mua/bán, khung tiêu đề, độ rộng
cột, chiều cao hàng và cố định khung theo mẫu VBA. File xuất là **.xlsx dữ liệu**,
gồm 8 sheet; không chứa macro, MENU hoặc Thamkhao.

| Sheet | Nội dung | Số cột |
|---|---|---:|
| TongHopHD_Mua / TongHopHD_Ban | Tổng hợp hóa đơn | 64 |
| ChiTietHD_Mua / ChiTietHD_Ban | Chi tiết hàng hóa | 34 |
| ChiTietHD_Mua_XML / ChiTietHD_Ban_XML | Chi tiết từ XML gốc | 32 |
| BaoCao_LoiTaiHD | Lỗi tải, parse và kết quả phục hồi | 17 |
| LinkTraCuu | Nhà cung cấp, link và nhãn trạng thái | 15 |

- **Ghi tiếp vào file đã chọn** mặc định bật: thêm hóa đơn mới, cập nhật hóa đơn
  trùng khóa, giữ các hóa đơn khác và bảng tra cứu đã sửa. CLI dùng APPEND_OUTPUT=true.
- XML lỗi HTTP 500/504: thử lấy API detail để bổ sung tổng hợp và chi tiết;
  hai sheet _XML chỉ chứa dữ liệu từ XML gốc.
- Ứng dụng lưu kết quả mặc định ở Documents\HDDT-Downloader; CLI dùng output/.
  Tên Excel mặc định là HoaDonDienTu.xlsx.

Xem [quy cách workbook](docs/EXCEL_FORMAT.md) và
[hướng dẫn ghi tiếp, link tra cứu, proxy, cập nhật](docs/USAGE.md).

## Chạy từ mã nguồn

Yêu cầu Windows PowerShell 5.1 hoặc PowerShell 7 trên Windows/macOS/Linux.
Sao chép [.env.example](.env.example) thành .env và điền:

```dotenv
GDT_USERNAME=tai_khoan_GDT
GDT_PASSWORD=mat_khau
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
INVOICE_DIRECTION=both
```

| Tác vụ | Windows | PowerShell 7 trên macOS/Linux |
|---|---|---|
| Tải từ GDT | run.cmd | pwsh -NoProfile -File Invoke-Hddt.ps1 |
| Nhập cấu hình trực tiếp | run-interactive.cmd | pwsh -NoProfile -File Invoke-Hddt.ps1 -Interactive |
| XML có sẵn → Excel | parse-local.cmd | pwsh -NoProfile -File Parse-LocalXml.ps1 |
| Giao diện Windows | run-gui.cmd | — |

Với XML offline, đặt LOCAL_XML_DIR trong .env. Toàn bộ cấu hình có trong
[.env.example](.env.example); cách dùng nâng cao ở [docs/USAGE.md](docs/USAGE.md).

## Phát triển

```powershell
# Kiểm thử lõi offline
powershell -NoProfile -File tests/Run-Tests.ps1
pwsh -NoProfile -File tests/Run-Tests.ps1

# Build EXE một file trên Windows
powershell -NoProfile -ExecutionPolicy Bypass -File Build-WindowsApp.ps1
```

[Kiến trúc](docs/ARCHITECTURE.md) · [Kiểm thử giao diện và phát hành](docs/USAGE.md#kiểm-thử)
· [Đóng góp](CONTRIBUTING.md) · [Hướng dẫn agent](AGENTS.md)

Mật khẩu và token không ghi vào log hoặc workbook. Không đưa .env hay hóa đơn
thật lên Git. Chi tiết: [SECURITY.md](SECURITY.md).
