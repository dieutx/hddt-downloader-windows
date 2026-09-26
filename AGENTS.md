# Hướng dẫn cho AI Coding Agent

File này dành cho Codex, Claude Code, Cursor, GitHub Copilot, OpenCode và các
AI Coding Agent khác làm việc trong repository `hddt-downloader-windows`.

## 1. Project là gì

Công cụ PowerShell tải **XML hóa đơn điện tử** từ hệ thống GDT
(`hoadondientu.gdt.gov.vn`) rồi xuất ra **workbook `.xlsx` data-only** theo
đúng bố cục sheet/cột của bản gốc `TaiHoaDonDienTu`.

- Không phải Excel add-in, **không có VBA**, không sửa template `.xlsm`.
- Chạy được trên **Windows PowerShell 5.1**, **PowerShell 7**, macOS và Linux.
- Không phụ thuộc thư viện ngoài: chỉ .NET (`System.IO.Compression`,
  `System.Xml`) và `System.Net.Http`/`Invoke-WebRequest`.

## 2. Trước khi sửa

Agent phải:

1. Đọc `README.md` (cách chạy, cấu hình, ý nghĩa sheet/cột).
2. Đọc `CONTRIBUTING.md` và file này.
3. Đọc `docs/ARCHITECTURE.md` để hiểu luồng dữ liệu.
4. Nếu đụng tới phần xuất Excel, đọc `docs/EXCEL_FORMAT.md` **bắt buộc**.
5. Hiểu behavior hiện tại trước khi đề xuất sửa.
6. Kiểm tra `git status` và bảo toàn thay đổi của người dùng (`.env`,
   `output/`, file `.xlsx` cục bộ đều là dữ liệu riêng của họ).

## 3. Bản đồ repository

| Đường dẫn | Vai trò |
|---|---|
| `Invoke-Hddt.ps1` | Entry point tải từ GDT: đăng nhập → danh sách → XML → Excel |
| `Parse-LocalXml.ps1` | Entry point parse XML có sẵn trong `LOCAL_XML_DIR` → Excel |
| `run.cmd`, `run-interactive.cmd` | Launcher Windows (`powershell.exe`) |
| `parse-local.cmd`, `parse-local-interactive.cmd` | Launcher Windows cho parse XML |
| `src/Logging.ps1` | `Write-HddtLog`, `Start-HddtLogging`, dừng an toàn (Ctrl+C) |
| `src/Config.ps1` | Đọc `.env`, kiểm tra giá trị, dựng object cấu hình |
| `src/Http.ps1` | `Invoke-GdtRequest`: token, proxy, retry, 429, log mạng |
| `src/Login.ps1` | CAPTCHA SVG → ký tự, đăng nhập, lấy token trong bộ nhớ |
| `src/InvoiceApi.ps1` | Gọi API GDT: danh sách hóa đơn, tải ZIP XML, chuỗi liên quan |
| `src/XmlParser.ps1` | Parse XML hóa đơn thành `Summary` + `Details` |
| `src/LinkTraCuu.ps1` | Bảng tra cứu (sheet `LinkTraCuu`) và nhãn `tthai`/`ttxly` |
| `src/ExcelExporter.ps1` | Dựng workbook `.xlsx` data-only (không thư viện ngoài) |
| `tests/Run-Tests.ps1` | Toàn bộ kiểm thử, chạy bằng một lệnh |
| `tests/fixtures/` | Fixture XML giả lập |

## 4. Kiểm thử

Bắt buộc chạy trước khi mở Pull Request:

```powershell
pwsh -NoProfile -File tests/Run-Tests.ps1      # PowerShell 7 (Windows/macOS/Linux)
powershell -NoProfile -File tests/Run-Tests.ps1  # Windows PowerShell 5.1
```

Script in `All tests passed.` khi đạt và **tự dừng với thông báo `FAILED:`** khi
lệch. Test bao gồm cả kiểm tra cấu trúc gói `.xlsx` sau khi xuất, nên đừng bỏ
qua khi sửa `src/ExcelExporter.ps1`.

Cách thêm test: thêm `Assert-Equal <kỳ vọng> <thực tế> '<mô tả>'` vào
`tests/Run-Tests.ps1`. Fixture phải là dữ liệu giả, không dùng MST thật.

## 5. Quy tắc bắt buộc

### Ngôn ngữ và tương thích

- PowerShell 5.1 và 7 đều phải chạy được: không dùng ternary (`? :`),
  toán tử `??`, `Join-Path` 3 tham số, `[ordered]` trong tham số bắt buộc, hay
  API chỉ có trong .NET Core.
- `Set-StrictMode -Version 2.0` ở đầu mỗi file `src/*.ps1`: mọi biến phải khởi
  tạo trước khi dùng, không dựa vào biến chưa khai báo.
- File lưu **UTF-8 không BOM**, xuống dòng **LF**.
- Comment và thông báo lỗi viết tiếng Việt không dấu nhất quán với file đang
  sửa; giữ nguyên cách diễn đạt hiện có.
- Không đổi tên hàm/biến public (`Export-InvoiceWorkbook`,
  `Invoke-GdtRequest`, `ConvertFrom-InvoiceXml`, `Get-HddtConfig`, …) nếu không
  cần thiết: chúng là hợp đồng giữa các file và giữa các lần chạy.

### Xuất Excel (xem `docs/EXCEL_FORMAT.md`)

- Workbook là gói OPC: **bắt buộc có** `[Content_Types].xml`, `_rels/.rels`,
  `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`.
- Đường dẫn part trong gói ghép **từng phần tử** qua
  `Join-ExcelPackagePath`, tuyệt đối không dùng chuỗi `'xl\worksheets\sheet1.xml'`.
  Trên macOS/Linux `\` thành dấu phân cách và `_rels\.rels` biến thành file ẩn,
  `Get-ChildItem` bỏ qua file ẩn nên gói thiếu quan hệ gốc và Excel báo
  *"we couldn't open your workbook"*.
- Danh sách part đóng gói là **tường minh** (`Add-ExcelPackagePart`), không quét
  thư mục tạm; `Assert-ExcelPackage` phải chạy trước khi thay thế file cũ.
- Thứ tự phần tử trong `worksheet.xml` phải theo schema SpreadsheetML
  (`sheetPr, dimension, sheetViews, sheetFormatPr, cols, sheetData, autoFilter,
  hyperlinks, printOptions, pageMargins, pageSetup`).
- Số kiểu dùng trong `Set-ExcelDataCell` là chỉ số `cellXfs` 0..16; thêm style
  mới thì sửa cả `Write-ExcelStylesXml`.
- Không thêm `vbaProject`, `MENU`, `Thamkhao` hay `tableParts`.

### Bảng tra cứu

- `src/LinkTraCuu.ps1` là **nguồn duy nhất** cho link tra cứu, tên trường mã tra
  cứu và nhãn `tthai`/`ttxly`. Thêm MST mới thì sửa bảng đó, đừng rải link
  vào `ExcelExporter.ps1`.
- Bảng lấy từ sheet `LinkTraCuu` của `TaiHoaDonDienTu`; một MST có nhiều dòng
  thì **dòng cuối thắng**, giữ đúng quy tắc của VBA gốc.

### Dữ liệu nhạy cảm (xem `SECURITY.md`)

- Không commit `.env`, token, cookie, `Authorization`, mật khẩu.
- Không commit XML/PDF/ZIP hóa đơn thật, MST thật, tên/địa chỉ khách hàng.
- Không log mật khẩu/token; `Protect-ExcelErrorText` phải redact trước khi ghi
  vào sheet lỗi.
- MST trong bảng `LinkTraCuu` là dữ liệu tham chiếu công khai (MST nhà cung cấp
  T-VAN) nên được giữ.

## 6. Không được làm

- Viết lại toàn bộ chương trình, đổi kiến trúc API hoặc endpoint khi chưa xác
  minh được.
- Thêm dependency ngoài (không có trình quản lý package trong repo).
- Tự tạo endpoint không có trong mã nguồn `TaiHoaDonDienTu`.
- Đổi luồng đăng nhập/CAPTCHA ngoài phạm vi task.
- Sửa hoặc xóa logic tương thích khi chưa hiểu và chưa có test hồi quy.
- Ghi đè `output/` của người dùng, `.env`, hoặc file `.xlsx` đã tải về.
- Thêm dữ liệu thật vào test/fixture.

## 7. Được khuyến khích

- Bug fix: parse JSON/XML, phân trang, retry 429, định dạng ngày/số, link tra
  cứu, đóng gói workbook.
- Khác biệt Windows/macOS/Linux (dấu phân cách đường dẫn, file ẩn, encoding,
  `Get-ChildItem`).
- Tương thích: GDT đổi response, schema XML mới, PowerShell 5.1 vs 7.
- Test hồi quy cho mọi bug đã sửa.
- Tài liệu: `README.md`, `docs/`, thông báo lỗi rõ ràng hơn.

## 8. Definition of Done

Một PR được coi là hoàn tất khi:

1. `tests/Run-Tests.ps1` chạy xanh (ghi rõ nếu không chạy được và lý do).
2. Behavior cũ được giữ hoặc thay đổi có nêu rõ trong PR và README.
3. Test mới cho bug vừa sửa (nếu có).
4. Không có file dữ liệu thật trong diff.
5. Commit rõ mục đích, không trộn thay đổi không liên quan.

## 9. Khi phát hiện vấn đề ngoài phạm vi

Không tự sửa lan sang. Ghi vào phần "Đề xuất" của PR hoặc mở Issue gồm: hiện
tượng, phạm vi, rủi ro, cách kiểm thử dự kiến.
