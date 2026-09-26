# Chính sách bảo mật

## Phạm vi dữ liệu

Công cụ làm việc với dữ liệu thật của người dùng:

- Tài khoản GDT (`GDT_USERNAME`, `GDT_PASSWORD`) trong file `.env`.
- Token đăng nhập chỉ tồn tại trong bộ nhớ của phiên chạy, không ghi ra đĩa.
- XML hóa đơn, workbook `.xlsx` và nhật ký trong `output/`.

## Không được commit

- `.env` và mọi bản sao chứa credential (`.env.*`, trừ `.env.example`).
- Token, cookie, JWT, header `Authorization`.
- XML/ZIP/PDF hóa đơn thật, MST của tài khoản, tên/địa chỉ khách hàng.
- File `.xlsx` đã xuất, thư mục `output/`, `logs/`.

`.gitignore` đã chặn sẵn các đường dẫn trên; nếu lỡ commit, hãy báo maintainer
để họ hỗ trợ gỡ khỏi lịch sử.

## Xử lý trong mã nguồn

- Mật khẩu và token không bao giờ được ghi vào log hoặc sheet `BaoCao_LoiTaiHD`;
  `Protect-ExcelErrorText` phải redact trước khi ghi.
- `Invoke-GdtRequest` gửi `Authorization` qua header, không nhúng vào URL.
- Credential proxy phải tách `PROXY_USERNAME`/`PROXY_PASSWORD`; URI có userinfo
  bị từ chối.
- Fixture kiểm thử phải là dữ liệu giả, không truy ngược được về người nộp thuế
  thật.

## Báo cáo lỗ hổng

Mở Issue kèm mô tả, bước tái hiện và mức độ ảnh hưởng. **Không** đính kèm
credential, token, hóa đơn thật hay workbook thật vào Issue. Nếu cần minh hoạ,
hãy dùng dữ liệu giả lập.

## Dữ liệu tham chiếu được phép giữ

Bảng `LinkTraCuu` trong `src/LinkTraCuu.ps1` chứa MST và địa chỉ trang tra cứu
của các nhà cung cấp dịch vụ hóa đơn điện tử (T-VAN). Đây là dữ liệu tham chiếu
công khai, cần giữ để tạo link tra cứu hóa đơn.
