# Định dạng workbook xuất ra

Tài liệu này là hợp đồng giữa chương trình và người dùng: tên sheet, tên cột, thứ
tự cột, kiểu số và cấu trúc gói `.xlsx`. Sửa phần xuất Excel mà không cập nhật tài
liệu này sẽ làm người dùng mất dữ liệu đọc được.

## Sheet

| # | Tên | Dòng tiêu đề | Bắt đầu dữ liệu | Số cột | Bộ lọc |
|---|---|---|---|---|---|
| 1 | `TongHopHD_Mua` | 2 (tiêu đề ở dòng 1) | 3 | 64 | không |
| 2 | `ChiTietHD_Mua` | 2 | 3 | 34 | không |
| 3 | `ChiTietHD_Mua_XML` | 2 | 3 | 32 | không |
| 4 | `TongHopHD_Ban` | 2 | 3 | 64 | không |
| 5 | `ChiTietHD_Ban` | 2 | 3 | 34 | không |
| 6 | `ChiTietHD_Ban_XML` | 2 | 3 | 32 | không |
| 7 | `BaoCao_LoiTaiHD` | 1 | 2 | 17 | có |
| 8 | `LinkTraCuu` | 1 | 2 | 15 (A..O) | không |

Mọi sheet trừ `BaoCao_LoiTaiHD` và `LinkTraCuu` đều có dòng 1 là tiêu đề lớn và
cố định hàng tiêu đề. Không có `vbaProject`, `MENU`, `Thamkhao`, `tableParts`.

## Sheet tổng hợp (1, 4)

| Cột | Nội dung | Ghi chú |
|---|---|---|
| 1 | STT | đánh số lại theo từng hướng |
| 2 | Tên HĐ | `tlhdon` |
| 3 | Mẫu HĐ | `khmshdon` |
| 4 | Ký hiệu hóa đơn | `khhdon` |
| 5 | Số hóa đơn | `shdon` |
| 6 | Ngày lập hóa đơn | kiểu ngày `dd/mm/yyyy` |
| 7 | Đơn vị tiền tệ | `dvtte` |
| 8 | Tỷ giá | số |
| 9-11 | Tên/MST/Địa chỉ người bán | `nbten`, `nbmst`, `nbdchi` |
| 12 | Ngày ký số người bán | kiểu ngày |
| 13-14 | Mã CQT, ngày cấp mã CQT | `mhdon`, `ncma` |
| 15-17 | Tên/MST/Địa chỉ người mua | `nmten`, `nmmst`, `nmdchi` |
| 18-41 | 6 nhóm thuế suất | mỗi nhóm 4 cột: thuế suất, thành tiền trước thuế, tiền thuế, giảm trừ thuế suất (`thttltsuat`) |
| 42 | Tổng tiền chưa thuế | `tgtcthue` |
| 43 | Tổng giảm trừ không chịu thuế | `tgtkcthue` |
| 44 | Tổng tiền thuế | `tgtthue` |
| 45-46 | Tên loại phí, tổng tiền phí | `thttlphi` |
| 47 | Tổng tiền chiết khấu thương mại | `ttcktmai` |
| 48 | Tổng giảm trừ khác | `tgtkhac` |
| 49-50 | Tổng thanh toán số, bằng chữ | `tgtttbso`, `tgtttbchu` |
| 51 | Ghi chú | `gchu` |
| 52 | Trạng thái hóa đơn | nhãn `tthai` (I2:I8 của `LinkTraCuu`) |
| 53 | Kết quả kiểm tra hóa đơn | nhãn `ttxly` (O2:O11, tra bằng chỉ số `ttxly + 1`) |
| 54 | MSTTCGP | `msttcgp` |
| 55 | Link tra cứu | xem "Link tra cứu" bên dưới |
| 56 | Mã tra cứu | `cttkhac`/`ttkhac`, hoặc mã riêng của VNPT/BKAV |
| 57 | Chuỗi hóa đơn liên quan | xuống dòng, tự giãn chiều cao |
| 58-63 | HĐ gốc: loại, mẫu, ký hiệu, số, ngày, ghi chú | chỉ ghi khi `tthai` 2-6 |
| 64 | Thông tin liên quan | xuống dòng, tự giãn chiều cao |

Cột 45/46 (`tgtkhac`, `tgtttbso`, `tgtttbchu`, `gchu`) giữ đúng vị trí của bản
gốc, không được dồn cột.

## Sheet chi tiết (2, 5)

Theo bản gốc v6.7.4, cột 1 là **Mẫu số hóa đơn** (`khmshdon`), rồi tới cột 2-15 là
thông tin chung của hóa đơn (ký hiệu, số, ngày, tiền tệ, tỷ giá, người bán 8-11,
mã CQT 11-12, người mua 13-15) và cuối cùng là từng dòng hàng hóa: số thứ tự 16,
tính chất, mã/tên hàng, DVT, số lượng, đơn giá, chiết khấu, loại thuế suất, thuế
suất 26, thành tiền chưa thuế 27, tiền thuế 28, thành tiền có thuế 29, tổng thuế
trên HĐ 30, chênh lệch kê khai thuế 31, MSTTCGP 32, link tra cứu 33, mã tra cứu 34.

Cột `Mẫu số hóa đơn` và hai cột MST người bán (8) / người mua (14) ghi dạng Text
(style 16) để giữ nguyên số `0` ở đầu.

## Sheet chi tiết XML (3, 6)

Cùng bố cục nhưng `Ngày lập hóa đơn` và `Thuế suất` giữ **dạng chữ** đúng như XML,
và sheet `ChiTietHD_Ban_XML` giữ `NBan`/`NMua` đúng theo tên node trong XML (không
đảo nhãn).

Dữ liệu lấy từ API detail sau XML HTTP 500 chỉ ghi vào sheet tổng hợp và chi
tiết (1, 2, 4, 5); không ghi vào hai sheet XML (3, 6). Trường API thiếu hoặc
null giữ trống nếu danh sách chưa có giá trị; số lượng/đơn giá không suy diễn.

## Sheet lỗi (7)

17 cột: STT, thời gian ghi nhận, loại hóa đơn, nguồn API, MST người bán, ký hiệu
mẫu số, ký hiệu hóa đơn, số hóa đơn, ngày lập, công đoạn lỗi, endpoint, HTTP
status, nội dung lỗi, số lần đã thử, `Retry-After`, kết quả cuối, ghi chú.

Nội dung lỗi đã redact qua `Protect-ExcelErrorText` (bỏ userinfo trong URL, `Bearer`,
`Authorization`, mật khẩu/proxy). Có `autoFilter`.

Nếu XML HTTP 500 được bổ sung dữ liệu bằng API detail, dòng báo cáo vẫn giữ
HTTP 500 và lỗi XML gốc; cột kết quả cuối ghi `Da lay du lieu tu API detail`,
cột ghi chú cho biết dữ liệu đã ghi vào tổng hợp/chi tiết và không có XML gốc.

## Sheet `LinkTraCuu` (8)

| Cột | Ý nghĩa |
|---|---|
| A | Tên tổ chức |
| B | MST nhà cung cấp (MSTTCGP) |
| C | MST người bán có link riêng |
| D | Link tra cứu |
| E | Tên trường trong `TTKhac` chứa mã tra cứu |
| F | Ghi chú |
| G | (trống) |
| H, I | `tthai` → Trạng thái hóa đơn |
| J | (trống) |
| K, L | `ttxly` → Kết quả kiểm tra (MUA) |
| M | (trống) |
| N, O | `ttxly` → Kết quả kiểm tra (BAN) |

Quy tắc:

- Một MST có nhiều dòng thì **dòng cuối thắng** (giữ đúng cách VBA gán đè vào
  dictionary).
- Đây là nguồn duy nhất cho link tra cứu và nhãn `tthai`/`ttxly`
  (`src/LinkTraCuu.ps1`).
- Người dùng sửa sheet này rồi trỏ `LOOKUP_TABLE_XLSX` vào file để dùng cho lần
  chạy sau.

## Link tra cứu (cột 55)

Thứ tự ưu tiên:

1. `MSTTCGP` từ danh sách GDT hoặc `TTChung/MSTTCGP` trong XML.
2. Nếu không có, tra MST người bán/người mua trong `LinkTraCuu` (ưu tiên link riêng
   theo cột C).
3. Không tra được thì ghi `Khong co link tra cuu` (sheet XML ghi
   `Khong tim thay link tra cuu`).

Với MSTTCGP có sẵn còn xử lý riêng, giữ như bản gốc: `0100684378` (VNPT) cần
MCCQT, `0101360697` (BKAV) cần `id` hóa đơn, `0105987432` dựng link từ MST người
bán. Chỉ URL HTTP/HTTPS tuyệt đối, có host và cú pháp hợp lệ mới được ghi kèm
hyperlink thật. URL mẫu chứa biểu thức ghép MST hoặc URL không hợp lệ vẫn giữ
nguyên nội dung ô dưới dạng chữ. Writer kiểm tra mọi hyperlink; validator gói
kiểm tra target của relationship trước khi thay thế file cũ.

## Kiểu số

Chỉ số truyền vào `Set-ExcelDataCell` là vị trí trong `cellXfs` của
`Write-ExcelStylesXml`:

| Chỉ số | Dùng cho |
|---|---|
| 0 | mặc định |
| 1 | dòng tiêu đề lớn |
| 2, 3, 4, 13 | dòng tiêu đề bảng (nền khác nhau, xuống dòng) |
| 5 | chữ (Calibri) |
| 6 | ngày `dd/mm/yyyy` |
| 7 | số, định dạng kế toán |
| 8 | phần trăm `0%` |
| 9 | chữ xuống dòng, canh trên |
| 10 | hyperlink (gạch chân, xanh) |
| 11, 12 | chữ xanh / chữ đỏ |
| 14 | ngày giờ `dd/mm/yyyy hh:mm:ss` |
| 15 | ngày |
| 16 | chữ, định dạng text |

Thêm style mới phải sửa cả `Write-ExcelStylesXml` và giữ `cellXfs count` khớp số
phần tử.

Toàn bộ font dùng Calibri; các chỉ số style, cỡ chữ, màu và định dạng số giữ
nguyên. Dữ liệu JSON UTF-8 được giải mã trước khi ghi XML của workbook.

## Gói `.xlsx` (OPC)

Workbook được dựng tay bằng `System.IO.Compression` và `System.Xml`, nên phải tuân
thủ đặc tả:

```text
[Content_Types].xml          (Default rels/xml + Override từng part)
_rels/.rels                  → officeDocument: xl/workbook.xml
xl/workbook.xml              (danh sách sheet, r:id khớp workbook.xml.rels)
xl/_rels/workbook.xml.rels   (7 worksheet + 1 styles)
xl/styles.xml                (numFmts, fonts, fills, borders, cellXfs, cellStyles)
xl/worksheets/sheet1..8.xml
xl/worksheets/_rels/sheetN.xml.rels   (chỉ khi sheet có hyperlink)
```

Quy tắc bắt buộc:

1. Đường dẫn part ghép **từng phần tử** bằng `Join-ExcelPackagePath`; không dùng
   chuỗi chứa `\`. Trên macOS/Linux `Join-Path` biến `\` thành dấu phân cách và
   `_rels\.rels` thành file ẩn `.rels`, mà `Get-ChildItem` mặc định bỏ qua file
   ẩn — gói thiếu quan hệ gốc và Excel báo *"we couldn't open your workbook"*.
2. Danh sách part đóng gói là tường minh (`Add-ExcelPackagePart`), không suy ra từ
   thư mục tạm.
3. `Assert-ExcelPackage` chạy trên file tạm **trước** khi thay thế file cũ; thiếu
   part bắt buộc thì ném lỗi và giữ nguyên file cũ.
4. Tên entry trong zip dùng `/`, không có ký tự đường dẫn.
5. Thứ tự phần tử trong `worksheet.xml`: `sheetPr, dimension, sheetViews,
   sheetFormatPr, cols, sheetData, autoFilter, hyperlinks, printOptions,
   pageMargins, pageSetup`.
6. Trong một `<row>`, các ô phải theo thứ tự cột tăng dần.
7. Text dùng `t="inlineStr"`; số và ngày dùng `t="n"` với ngày ở dạng OADate.
8. `workbook.xml.rels` phải có đúng một styles relationship và đúng số worksheet.

## Kiểm tra gói sau khi đổi

```powershell
pwsh -NoProfile -File tests/Run-Tests.ps1
```

Test mở lại file `.xlsx` vừa xuất và kiểm tra: đủ part bắt buộc, không entry nào
giữ dấu `\`, quan hệ gốc trỏ đúng `xl/workbook.xml`, số sheet, thứ tự tên sheet,
quan hệ hyperlink, và thứ tự phần tử `font` theo schema.
