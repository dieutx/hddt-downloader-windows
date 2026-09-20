# HDDT Downloader for Windows

Công cụ PowerShell thuần Windows để tải XML hóa đơn điện tử từ API GDT, parse dữ liệu và tạo file Excel `.xlsx`. Không cần cài Python, Node.js, 7-Zip, NuGet, PowerShell Gallery hoặc Microsoft Excel.

## Chức năng

- Tải hóa đơn mua vào, bán ra hoặc cả hai.
- Hỗ trợ nguồn `query` và `sco-query`.
- Chỉ lưu XML vào một thư mục chung `output\xml`; dữ liệu ZIP được xử lý trong RAM.
- Tạo workbook gồm bảng tổng hợp, chi tiết và danh sách lỗi.
- Parse riêng một thư mục XML có sẵn.
- Log UTF-8, retry lỗi mạng và tự giảm tốc khi gặp HTTP 429.

## Yêu cầu

- Windows 10/11 hoặc Windows Server.
- Windows PowerShell 5.1.
- Bearer token hợp lệ từ `hoadondientu.gdt.gov.vn`.

## Cách chạy

1. Sao chép `.env.example` thành `.env` nếu chưa có.
2. Điền `GDT_TOKEN`, khoảng ngày và loại hóa đơn.
3. Đóng file Excel đầu ra nếu đang mở.
4. Chạy `run.cmd`.

Kết quả mặc định:

```text
output/
  HoaDonDienTu.xlsx
  xml/
    purchase_query_MST_MAU_KYHIEU_SO.xml
  logs/
    download_YYYYMMDD_HHMMSS.log
```

API trả từng hóa đơn dưới dạng ZIP. Chương trình giải nén trong RAM và chỉ ghi XML xuống ổ đĩa. Nếu một phản hồi có nhiều XML, tên file được thêm hậu tố `_1`, `_2`, v.v.

## Parse XML local

Điền các biến sau rồi chạy `parse-local.cmd`:

```dotenv
LOCAL_XML_DIR=E:\duong-dan\toi\xml
LOCAL_DIRECTION=purchase
LOCAL_OUTPUT_XLSX=HoaDonDienTu_Local.xlsx
```

Chế độ này không gọi API và không cần token.

## Cấu hình chính

| Biến | Ý nghĩa |
|---|---|
| `GDT_TOKEN` | Token thuần hoặc chuỗi `Bearer ...` |
| `INVOICE_DIRECTION` | `purchase`, `sold` hoặc `both` |
| `FROM_DATE`, `TO_DATE` | Khoảng ngày dạng `dd/MM/yyyy` |
| `INCLUDE_REGULAR` | Lấy nguồn `query` |
| `INCLUDE_SCO` | Lấy nguồn `sco-query` |
| `OUTPUT_DIR` | Thư mục đầu ra |
| `OUTPUT_XLSX` | Tên workbook |
| `OVERWRITE_OUTPUT` | Cho phép ghi đè workbook |
| `REDOWNLOAD_XML` | `true`: luôn tải lại; `false`: dùng XML đã có |
| `REQUEST_DELAY_MS` | Khoảng nghỉ giữa request; khuyến nghị `600` |
| `ADAPTIVE_THROTTLE` | Tự giảm tốc khi gặp HTTP 429 |
| `MAX_RETRIES` | Số lần retry lỗi mạng/429/5xx |
| `HTTP_TIMEOUT_SECONDS` | Timeout của mỗi request |
| `LOG_LEVEL` | `debug`, `info`, `warn` hoặc `error` |
| `LOG_TO_FILE` | Lưu log trong `OUTPUT_DIR\logs` |
| `PROGRESS_EVERY` | In tiến độ sau mỗi N hóa đơn |

URL API được cố định trong code. Authorization header không được gửi tới host khác. Token không được ghi vào log hoặc workbook.

## Tốc độ và HTTP 429

Benchmark thực tế với 224 hóa đơn:

- Tuần tự, `REQUEST_DELAY_MS=600`: khoảng 2 phút 31 giây, không gặp 429.
- `REQUEST_DELAY_MS=500`: gặp 429 sau khoảng 145 hóa đơn.
- Nhiều luồng: không nhanh hơn do GDT giới hạn chung theo token/IP.

Vì vậy repo sử dụng thuật toán tuần tự 600 ms và tái sử dụng kết nối HTTP. Nếu vẫn gặp 429, dừng 2–5 phút rồi tăng `REQUEST_DELAY_MS` lên `1000` hoặc `1500`.

HTTP 500 từ `export-xml` thường có nghĩa hóa đơn không có hồ sơ XML gốc. Chương trình vẫn gọi thử từng hóa đơn, ghi lỗi vào workbook và tiếp tục hóa đơn kế tiếp.

## Kiểm thử offline

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
```

Test không gọi mạng và không cần token thật.

## An toàn khi push Git

`.gitignore` loại trừ:

- `.env` và token thật;
- `output/`, log và XML tải về;
- file `.xlsx`, `.xlsm`, `.zip`;
- thư mục tạm `work/`.

Chỉ tải dữ liệu mà tài khoản của bạn được phép truy cập.
