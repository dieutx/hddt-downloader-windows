# HDDT Downloader for Windows

Công cụ PowerShell thuần Windows để tải XML hóa đơn điện tử từ GDT và xuất dữ liệu ra Excel `.xlsx`. Không cần cài Python, Node.js, thư viện PowerShell hay Microsoft Excel.

Repo có hai chức năng độc lập:

1. Tải XML từ GDT, sau đó parse và tạo file Excel.
2. Chỉ parse một thư mục XML đã có sẵn, không gọi GDT và không cần token.

## Cài đặt từ GitHub

Mở **Command Prompt (CMD)** và kiểm tra Git:

```bat
git --version
```

Nếu CMD báo không tìm thấy lệnh `git`, cài Git bằng công cụ có sẵn trên Windows 10/11:

```bat
winget install --id Git.Git -e --source winget
```

Sau khi cài, đóng CMD, mở lại rồi kiểm tra `git --version`. Nếu máy không có `winget`, tải Git for Windows tại [git-scm.com/download/win](https://git-scm.com/download/win).

Clone và tạo file cấu hình:

```bat
git clone https://github.com/dieutx/hddt-downloader-windows.git
cd hddt-downloader-windows
copy .env.example .env
notepad .env
```

## Chức năng 1: Tải XML từ GDT và tạo Excel

### Lấy token đăng nhập GDT

1. Đăng nhập tại [hoadondientu.gdt.gov.vn](https://hoadondientu.gdt.gov.vn/).
2. Khi vẫn đang ở trang GDT, nhấn `F12` và mở tab **Console**.
3. Dán đoạn JavaScript sau rồi nhấn `Enter`:

```javascript
const t = (document.cookie.match(/eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/) || [])[0];
t ? console.log("Bearer " + t) : console.log("❌ Không tìm thấy token");
```

4. Sao chép chuỗi bắt đầu bằng `Bearer ` và điền vào `GDT_TOKEN` trong `.env`:

```dotenv
GDT_TOKEN=Bearer eyJ...
```

Token là thông tin đăng nhập nhạy cảm: chỉ chạy đoạn mã trên đúng website GDT, không gửi token cho người khác và không đưa token vào Git. Nếu báo không tìm thấy token, hãy đăng nhập lại hoặc tải lại trang GDT rồi thử lại.

Điền tối thiểu các dòng sau trong `.env`:

```dotenv
GDT_TOKEN=token_cua_ban
INVOICE_DIRECTION=purchase
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
INCLUDE_REGULAR=true
INCLUDE_SCO=true
REDOWNLOAD_XML=true
```

Ý nghĩa bốn tùy chọn dễ nhầm:

- `INVOICE_DIRECTION=purchase`: lấy hóa đơn mua vào, tức hóa đơn nhà cung cấp xuất cho đơn vị đang tra cứu.
- `INVOICE_DIRECTION=sold`: lấy hóa đơn bán ra, tức hóa đơn do đơn vị đang tra cứu phát hành.
- `INVOICE_DIRECTION=both`: lấy cả hóa đơn mua vào và bán ra.
- `INCLUDE_REGULAR=true`: lấy hóa đơn điện tử thông thường trên hệ thống GDT. Trong API, nguồn này có tên `query`.
- `INCLUDE_SCO=true`: lấy hóa đơn điện tử khởi tạo từ máy tính tiền. Trong API, nguồn này có tên `sco-query`.
- `REDOWNLOAD_XML=true`: mỗi lần chạy đều tải lại toàn bộ XML từ GDT và ghi lại file cùng tên nếu đã tồn tại.
- `REDOWNLOAD_XML=false`: dùng lại XML cùng tên đã có trong thư mục đầu ra và chỉ tải những file còn thiếu.

Đóng file Excel đầu ra nếu đang mở, sau đó chạy:

```bat
run.cmd
```

Chương trình tải từng hóa đơn và tạo workbook gồm bảng tổng hợp, bảng chi tiết và danh sách lỗi. XML mua vào và bán ra được tách thành hai thư mục; không tạo thư mục riêng cho từng hóa đơn. Dữ liệu ZIP từ API chỉ được giải nén trong bộ nhớ, không lưu thành file `.zip`.

Kết quả mặc định:

```text
output\
  HoaDonDienTu.xlsx
  xml\
    purchase\
      purchase_query_MST_MAU_KYHIEU_SO.xml
    sold\
      sold_query_MST_MAU_KYHIEU_SO.xml
  logs\
    download_YYYYMMDD_HHMMSS.log
```

## Chức năng 2: Chỉ parse thư mục XML có sẵn

Chức năng này không tải dữ liệu, không gọi API GDT và không sử dụng `GDT_TOKEN`. Chương trình đọc các file `.xml` trong `LOCAL_XML_DIR` và các thư mục con, sau đó tạo một file Excel mới.

Điền trong `.env`:

```dotenv
LOCAL_XML_DIR=output\xml
LOCAL_DIRECTION=auto
LOCAL_OUTPUT_XLSX=HoaDonDienTu_Local.xlsx
```

Sau đó chạy:

```bat
parse-local.cmd
```

`LOCAL_DIRECTION` quyết định cách phân loại XML:

- `auto`: tự nhận diện từng file theo thư mục `purchase`/`sold` hoặc tiền tố tên file. Dùng lựa chọn này khi parse toàn bộ `output\xml`.
- `purchase`: ép toàn bộ XML trong thư mục thành hóa đơn mua vào.
- `sold`: ép toàn bộ XML trong thư mục thành hóa đơn bán ra.

Vì vậy có thể đặt `LOCAL_XML_DIR=output\xml` để parse cả hai nhóm trong một lần mà không bị gắn sai loại.

File Excel được tạo trong `OUTPUT_DIR`; mặc định là `output\HoaDonDienTu_Local.xlsx`.

## Toàn bộ tham số trong `.env`

### Dùng khi tải từ GDT

| Biến | Giá trị có thể nhập | Ý nghĩa |
|---|---|---|
| `GDT_TOKEN` | Token thuần hoặc `Bearer ...` | Token đăng nhập GDT; chỉ `run.cmd` sử dụng. |
| `INVOICE_DIRECTION` | `purchase`, `sold`, `both` | Chọn mua vào, bán ra hoặc cả hai. |
| `FROM_DATE` | `dd/MM/yyyy` | Ngày bắt đầu của kỳ hóa đơn. |
| `TO_DATE` | `dd/MM/yyyy` | Ngày kết thúc của kỳ hóa đơn. |
| `INCLUDE_REGULAR` | `true`, `false` | Bật/tắt hóa đơn điện tử thông thường (`query`). |
| `INCLUDE_SCO` | `true`, `false` | Bật/tắt hóa đơn khởi tạo từ máy tính tiền (`sco-query`). |
| `REDOWNLOAD_XML` | `true`, `false` | `true`: luôn tải lại tất cả XML; `false`: dùng lại XML cùng tên đã có. |
| `PAGE_SIZE` | Số nguyên dương | Số hóa đơn yêu cầu trong mỗi trang danh sách. |
| `REQUEST_DELAY_MS` | Số mili giây, ví dụ `600` | Khoảng nghỉ cơ bản giữa hai lần gọi API. |
| `ADAPTIVE_THROTTLE` | `true`, `false` | Cho phép tự điều chỉnh khoảng nghỉ theo phản hồi của máy chủ. |
| `MAX_RETRIES` | Số nguyên từ `0` trở lên | Số lần thử lại khi yêu cầu tạm thời thất bại. |
| `HTTP_TIMEOUT_SECONDS` | Số giây | Thời gian chờ tối đa cho một yêu cầu HTTP. |

### Đầu ra và log

| Biến | Giá trị có thể nhập | Ý nghĩa |
|---|---|---|
| `OUTPUT_DIR` | Đường dẫn tương đối hoặc tuyệt đối | Thư mục chứa XML, Excel và log. |
| `OUTPUT_XLSX` | Tên file `.xlsx` | File Excel của chức năng tải và parse. |
| `OVERWRITE_OUTPUT` | `true`, `false` | Cho phép ghi đè file Excel cùng tên. |
| `LOG_LEVEL` | `debug`, `info`, `warn`, `error` | Mức chi tiết của log. |
| `LOG_TO_FILE` | `true`, `false` | Bật/tắt lưu log thành file. |
| `PROGRESS_EVERY` | Số nguyên dương | In một dòng tiến độ sau mỗi N hóa đơn. |

### Dùng khi chỉ parse XML có sẵn

| Biến | Giá trị có thể nhập | Ý nghĩa |
|---|---|---|
| `LOCAL_XML_DIR` | Đường dẫn thư mục | Thư mục chứa XML cần parse; có đọc cả thư mục con. |
| `LOCAL_DIRECTION` | `auto`, `purchase`, `sold` | `auto` nhận diện từng file; hai giá trị còn lại ép toàn bộ thư mục về một loại. |
| `LOCAL_OUTPUT_XLSX` | Tên file `.xlsx` | Tên file Excel được tạo từ XML có sẵn. |

URL API được cố định trong mã nguồn. Authorization header không được gửi tới host khác. Token không được ghi vào log hoặc workbook.

## Kiểm thử offline

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
```

Test không gọi mạng và không cần token thật.
