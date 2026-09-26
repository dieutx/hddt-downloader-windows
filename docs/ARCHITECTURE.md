# Kiến trúc

## Tổng quan

```text
.env ──► Config.ps1 ──► cấu hình (token rỗng, hướng, kỳ, output)
                             │
        ┌────────────────────┴─────────────────────┐
        │ Invoke-Hddt.ps1 (có mạng)                 │ Parse-LocalXml.ps1 (XML sẵn có)
        │  Login.ps1  → CAPTCHA + token            │  XmlParser.ps1
        │  Http.ps1   → request/retry/proxy        │
        │  InvoiceApi → danh sách + tải ZIP XML    │
        │  XmlParser  → Summary + Details          │
        └────────────────────┬─────────────────────┘
                             ▼
                    ExcelExporter.ps1  ──►  output/HoaDonDienTu.xlsx
                    LinkTraCuu.ps1     ──►  bảng tra cứu + sheet LinkTraCuu
```

Cả hai entry point kết thúc ở `Export-InvoiceWorkbook`, nên mọi thay đổi về
xuất Excel chỉ cần làm trong `src/ExcelExporter.ps1` và
`src/LinkTraCuu.ps1`.

## Module

| Module | Trách nhiệm | Ghi chú |
|---|---|---|
| `Logging.ps1` | Log có mức độ, ghi file UTF-8, cờ dừng an toàn (Ctrl+C) | Không phụ thuộc module khác |
| `Config.ps1` | Đọc `.env`, chuẩn hoá và kiểm tra giá trị, dựng object cấu hình | `Get-HddtConfig` là hợp đồng của cả hai entry point |
| `Http.ps1` | `Invoke-GdtRequest`: thêm `Authorization`, proxy, retry, backoff 429, log mạng | 429 có lịch riêng, không phụ thuộc `MAX_RETRIES` |
| `Login.ps1` | Đọc CAPTCHA SVG, giải mã thành ký tự, đăng nhập, trả token | Token chỉ trong bộ nhớ |
| `InvoiceApi.ps1` | Danh sách hóa đơn (phân trang theo `state`), tải ZIP XML, chuỗi hóa đơn liên quan | Giữ nguyên item JSON gốc trong `GdtIndex` |
| `XmlParser.ps1` | XML hóa đơn → `Summary` + `Details` | Đọc `local-name()` nên không phụ thuộc namespace |
| `Parallel.ps1` | Tải + parse XML đa luồng bằng runspace pool, tự điều tiết số luồng theo tải máy chủ | Chỉ dùng khi `DOWNLOAD_WORKERS > 1`; mỗi hóa đơn xử lý đúng một lần theo chỉ số dùng chung |
| `LinkTraCuu.ps1` | Bảng tra cứu, tên trường mã tra cứu, nhãn `tthai`/`ttxly`, nạp bảng của người dùng | Nguồn dữ liệu tham chiếu duy nhất |
| `ExcelExporter.ps1` | Dựng workbook `.xlsx` data-only, đóng gói OPC, kiểm tra gói | Không dùng thư viện ngoài |

Các module được **dot-source** (`Import-Module` không dùng), nên hàm và biến
`$script:` dùng chung giữa các file trong cùng một phiên script.

## Luồng dữ liệu

### 1. Cấu hình

`Get-HddtConfig` đọc `.env` (hoặc hỏi trực tiếp với `-Interactive`), kiểm tra
khoảng giá trị rồi trả về một `pscustomobject`. `Token` luôn rỗng lúc khởi tạo vì
đăng nhập diễn ra sau đó.

### 2. Đăng nhập

`Invoke-GdtLogin` lấy CAPTCHA dạng SVG, `ConvertFrom-SvgCaptcha` giải mã các
lệnh vẽ thành ký tự (đọc thứ tự theo tọa độ `x`), rồi đăng nhập. CAPTCHA sai thì
thử lại tối đa 3 lần. Lỗi 401 không bị retry mù.

### 3. Danh sách hóa đơn

`Get-GdtInvoiceIndex` gọi `/api/{query|sco-query}/invoices/{direction}` theo từng
kỳ tháng, phân trang bằng tham số `state`. Mỗi item giữ nguyên JSON gốc ở
`GdtIndex` để phần xuất Excel đọc được mọi trường (kể cả `msttcgp`, `cttkhac`,
`ttkhac`) mà không cần chốt danh sách trường.

Lỗi một kỳ không làm mất các kỳ còn lại: lỗi được ghi vào `GdtIndexErrors` và
xuất ra sheet `BaoCao_LoiTaiHD`.

### 4. Tải và parse XML

`Save-GdtInvoiceXml` tải ZIP, chỉ giữ file `.xml`, ghi ra
`output/xml/<direction>/`. `REDOWNLOAD_XML=false` thì tận dụng lại file đã có
(kiểm tra tái sử dụng lọc theo tiền tố tên file để không quét toàn thư mục).

Khi `DOWNLOAD_WORKERS > 1`, giai đoạn tải + parse XML chạy song song trên một
runspace pool: mỗi hóa đơn lấy từ một chỉ số dùng chung nên xử lý đúng một lần,
kết quả kèm `Index` rồi được gom lại theo đúng thứ tự danh sách (không trùng,
không thiếu dòng). Giãn cách request, mốc tạm dừng khi HTTP 429 và số luồng hiện
hành dùng chung qua một "gate" (`New-HddtSharedGate` trong `src/Http.ps1`): gặp
429 thì giảm số luồng và tạm dừng toàn cục, khi ổn định lại thì tăng dần tới
trần `DOWNLOAD_WORKERS`. Danh sách (phân trang theo `state`) và phần hóa đơn
liên quan vẫn tuần tự.
`ConvertFrom-InvoiceXml` trả về:

```text
Summary : Direction, Source, XmlFile, InvoiceId, TemplateCode, InvoiceSeries,
          InvoiceNumber, InvoiceDate, Seller*, Buyer*, TaxAuthority*,
          ProviderTaxCode (TTChung/MSTTCGP), TaxAmount, InvoiceLookupFields,
          InvoiceAdditionalFields, ...
Details : mỗi dòng hàng hóa (LineNumber, Description, Quantity, UnitPrice,
          TaxRate, TaxAmount, AmountBeforeTax, AmountWithTax, ...)
```

`Merge-HddtParsedSummary` trong `Invoke-Hddt.ps1` bổ sung dữ liệu XML vào dòng
tổng hợp đã tạo từ danh sách, nhưng **không** ghi đè các trường đã có
(`GdtIndex`, `Status`, `RelatedChain`, …). Nhờ vậy hóa đơn tải XML lỗi vẫn còn
trong tổng hợp kèm dòng lỗi.

### 5. Hóa đơn liên quan

Hóa đơn trạng thái 2-6 gọi thêm `relative` và `related`, kết quả ghi thẳng vào
các cột 57-64 của sheet tổng hợp. `FETCH_RELATED=false` chỉ lấy dữ liệu có
sẵn trong danh sách để không tốn request.

### 6. Xuất Excel

`Export-InvoiceWorkbook` nhận ba mảng dữ liệu thô (`SummaryRows`, `DetailRows`,
`ErrorRows`) và tự chia theo hướng, dựng 8 sheet, đóng gói OPC rồi mới thay thế
file cũ. Hàm trả về thống kê (số dòng, số link tra cứu tìm được) để entry point
ghi log.

Chi tiết về workbook: `EXCEL_FORMAT.md`.

## Điểm mở rộng

| Cần làm | Chỗ sửa |
|---|---|
| Thêm nhà cung cấp T-VAN | thêm dòng vào bảng trong `src/LinkTraCuu.ps1` |
| Đổi câu chữ khi không có link | `Get-ExcelLookupLink` trong `src/ExcelExporter.ps1` |
| Thêm cột vào sheet tổng hợp | `$script:SourceSummaryHeaders` + `$script:SourceSummaryWidths` + chỉ số cột trong `New-ExcelSummaryRows` (giữ nguyên thứ tự cột cũ) |
| Thêm style | `Write-ExcelStylesXml` rồi dùng chỉ số `cellXfs` mới |
| Thêm biến cấu hình | `src/Config.ps1` + `.env.example` + `README.md` + test trong `tests/Run-Tests.ps1` |
| Hỗ trợ nguồn hóa đơn mới | thêm hàm tương tự `Get-GdtInvoiceIndex`, trả về cùng cấu trúc `GdtIndex` |

## Bất biến cần giữ

1. Không ghi credential ra log, file hay workbook.
2. Lỗi một phần (một kỳ, một hóa đơn, một XML) không được làm mất phần đã tải.
3. Xuất Excel là thao tác nguyên tử: ghi file tạm, kiểm tra gói, rồi mới thay
   thế; file cũ phải còn nguyên nếu bất kỳ bước nào thất bại.
4. Dữ liệu xuất ra phải giữ đúng tên sheet, tên cột, thứ tự cột và kiểu số của
   bản gốc `TaiHoaDonDienTu`.
5. Mọi thay đổi phải kèm test chạy được bằng `tests/Run-Tests.ps1`.
