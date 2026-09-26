# HDDT Downloader

Công cụ PowerShell để tải **XML hóa đơn điện tử** từ GDT và xuất thành file Excel.

- Tải hóa đơn **mua vào** và **bán ra**.
- Tự động lấy CAPTCHA và đăng nhập bằng tài khoản GDT.
- Có thể dùng proxy/VPN khi IP hiện tại bị GDT hạn chế.
- Xuất Excel dạng dữ liệu, không chứa VBA hoặc sheet `MENU`.
- Có thể chạy bằng `.env` hoặc nhập cấu hình trực tiếp.
<img width="1044" height="560" alt="image" src="https://github.com/user-attachments/assets/c0268a28-88ca-4e80-9df8-a0aa984faaa1" />

---

## 1. Chạy nhanh trên Windows

### Bước 1 — Tải mã nguồn

```bat
git clone https://github.com/dieutx/hddt-downloader-windows.git
cd hddt-downloader-windows
copy .env.example .env
```

### Bước 2 — Điền thông tin

Mở `.env` và sửa các dòng cần thiết:

```dotenv
GDT_USERNAME=0123456789
GDT_PASSWORD=mat_khau
FROM_DATE=01/09/2026
TO_DATE=30/09/2026
INVOICE_DIRECTION=both
```

| Biến | Ý nghĩa |
|---|---|
| `GDT_USERNAME` | Mã số thuế tài khoản GDT |
| `GDT_PASSWORD` | Mật khẩu GDT |
| `FROM_DATE`, `TO_DATE` | Khoảng ngày lập hóa đơn, định dạng `dd/MM/yyyy` |
| `INVOICE_DIRECTION` | `purchase` = mua vào, `sold` = bán ra, `both` = cả hai |

### Bước 3 — Chạy

Nhấn đúp:

```text
run.cmd
```

Chương trình tự lấy CAPTCHA, đăng nhập, tải danh sách hóa đơn, tải XML và tạo Excel.

---

## 2. Chế độ nhập trực tiếp

Nếu không muốn chỉnh `.env`, chạy:

```text
run-interactive.cmd
```

Chương trình sẽ hỏi:

1. Tài khoản và mật khẩu GDT.
2. Khoảng ngày và chiều hóa đơn.
3. Thư mục/file Excel đầu ra.
4. Có dùng proxy hay không.

Mật khẩu được nhập ở chế độ bí mật, không hiển thị trên màn hình và không ghi đè `.env`.
Nhập `CLEAR` ở ô proxy nếu muốn tắt proxy.

---

## 3. Chạy trên Linux / macOS

Cần PowerShell 7 trở lên.

### macOS — bản preview

```bash
brew install --cask powershell@preview
pwsh-preview --version
```

Bản preview dùng lệnh **`pwsh-preview`**.

### Ubuntu / Debian

Cài PowerShell 7 ổn định từ kho Microsoft:

```bash
sudo apt-get update
sudo apt-get install -y wget apt-transport-https software-properties-common

. /etc/os-release
wget -q "https://packages.microsoft.com/config/$ID/$VERSION_ID/packages-microsoft-prod.deb"
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

sudo apt-get update
sudo apt-get install -y powershell
pwsh --version
```

Lệnh `pwsh --version` phải trả về số phiên bản PowerShell 7 trở lên. Nếu dùng bản preview trên Linux, lệnh tương ứng là `pwsh-preview --version`.

Các bản phân phối khác: xem [hướng dẫn cài PowerShell trên Linux](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-linux).

### Chạy chương trình

```bash
# Linux hoặc bản PowerShell ổn định
pwsh -NoProfile -File Invoke-Hddt.ps1
pwsh -NoProfile -File Invoke-Hddt.ps1 -Interactive

# macOS bản preview
pwsh-preview -NoProfile -File Invoke-Hddt.ps1
pwsh-preview -NoProfile -File Invoke-Hddt.ps1 -Interactive
```

---

## 4. Dùng proxy

Nếu GDT không truy cập được từ IP hiện tại, thêm vào `.env`:

```dotenv
PROXY_URL=http://127.0.0.1:8080
PROXY_USERNAME=
PROXY_PASSWORD=
```

Proxy cần đăng nhập:

```dotenv
PROXY_URL=http://proxy.example:8080
PROXY_USERNAME=proxy-user
PROXY_PASSWORD=proxy-password
```

Cũng có thể dùng dạng rút gọn trong `PROXY_URL`:

```text
host:port:username:password
```

Nên tách username/password vào hai biến riêng để không làm lộ mật khẩu trong URL. Proxy được áp dụng cho đăng nhập, danh sách, related và tải XML.

---

## 5. Kết quả xuất

File mặc định nằm trong `output/`:

```text
output/
  HoaDonDienTu.xlsx
  xml/
    purchase/
    sold/
  logs/
```

Workbook gồm 8 sheet:

| Sheet | Nội dung |
|---|---|
| `TongHopHD_Mua` | Tổng hợp hóa đơn mua vào |
| `ChiTietHD_Mua` | Chi tiết dữ liệu mua vào |
| `ChiTietHD_Mua_XML` | Chi tiết đọc trực tiếp từ XML mua vào |
| `TongHopHD_Ban` | Tổng hợp hóa đơn bán ra |
| `ChiTietHD_Ban` | Chi tiết dữ liệu bán ra |
| `ChiTietHD_Ban_XML` | Chi tiết đọc trực tiếp từ XML bán ra |
| `BaoCao_LoiTaiHD` | Các lỗi tải/parse và kết quả xử lý |
| `LinkTraCuu` | Bảng định tuyến link tra cứu hóa đơn |

File Excel không có VBA, `MENU` hay `Thamkhao`.

### Link tra cứu hóa đơn

Cột **Link tra cứu** (cột 55) và **Mã tra cứu** (cột 56) của `TongHopHD_Mua`/`TongHopHD_Ban` được sinh từ sheet `LinkTraCuu`:

| Cột | Ý nghĩa |
|---|---|
| A | Tên tổ chức |
| B | MST nhà cung cấp dịch vụ T-VAN (MSTTCGP) |
| C | MST người bán cần link riêng (ưu tiên hơn link chung) |
| D | Link tra cứu |
| E | Tên trường trong `TTKhac` chứa mã tra cứu |
| F | Ghi chú |

Thứ tự ưu tiên khi sinh link:

1. `MSTTCGP` do GDT trả về, hoặc `TTChung/MSTTCGP` trong XML.
2. Nếu không có, tra MST của người bán/người mua trong bảng `LinkTraCuu`: hóa đơn phát hành trực tiếp qua MSSVĐHĐN thường không có MSTTCGP, nhưng bên kia vẫn có thể là một nhà cung cấp có trong bảng.
3. Không tra được thì ghi `Khong co link tra cuu`, đúng như bản gốc.

Muốn tự thêm/sửa link (nhà cung cấp mới, hoặc MST chi nhánh), mở file Excel đã xuất, sửa sheet `LinkTraCuu`, lưu lại rồi trỏ biến sau vào file đó:

```dotenv
LOOKUP_TABLE_XLSX=output/HoaDonDienTu.xlsx
```

Các dòng trong bảng đó được nối vào bảng gốc (dòng cuối thắng) rồi ghi lại vào workbook của lần chạy tiếp theo, nên sửa bao nhiêu lần cũng không nhân bản bảng.

---

## 6. Parse XML có sẵn

Không cần tài khoản hoặc Internet khi đã có sẵn file XML.

Đặt `LOCAL_XML_DIR` trong `.env`, sau đó chạy:

```bat
parse-local.cmd
```

Hoặc nhập trực tiếp:

```bat
parse-local-interactive.cmd
```

Trên Linux/macOS dùng `pwsh -NoProfile -File Parse-LocalXml.ps1` (hoặc `pwsh-preview` trên macOS preview).

---

## 7. Một số tuỳ chọn thường dùng

| Biến | Ý nghĩa |
|---|---|
| `INCLUDE_REGULAR` | Lấy hóa đơn thông thường |
| `INCLUDE_SCO` | Lấy hóa đơn máy tính tiền |
| `FETCH_RELATED` | Lấy chuỗi hóa đơn thay thế/điều chỉnh |
| `REDOWNLOAD_XML` | Tái sử dụng XML đã tải hoặc tải lại |
| `OVERWRITE_OUTPUT` | Ghi đè file Excel đã có |
| `LOOKUP_TABLE_XLSX` | Workbook `.xlsx` có sheet `LinkTraCuu` dùng để tra link tra cứu |
| `REQUEST_DELAY_MS` | Nghỉ giữa các request |
| `DOWNLOAD_WORKERS` | Số luồng tải XML tối đa (1-16; mặc định 3). Chương trình tự giảm khi GDT quá tải rồi tăng lại khi ổn định. Đặt 1 để tải tuần tự. |
| `MAX_RETRIES` | Số lần thử lại lỗi mạng/5xx |
| `HTTP_TIMEOUT_SECONDS` | Thời gian chờ một request |

Xem toàn bộ tuỳ chọn trong `.env.example`.

---

## 8. Lỗi và tạm dừng

- Lỗi HTTP 500/504 hoặc lỗi parse được ghi vào `BaoCao_LoiTaiHD`; chương trình tiếp tục xử lý phần còn lại.
- Nếu đăng nhập báo HTTP 401, CAPTCHA đã được đọc; hãy kiểm tra `GDT_USERNAME`, `GDT_PASSWORD`, quyền truy cập tài khoản và proxy trong `.env`.
- Nhấn `Ctrl+C` một lần để dừng sau request hiện tại; dữ liệu đã tải vẫn được xuất Excel.
- Nếu không lấy được hóa đơn nào, workbook lỗi vẫn được tạo và chương trình trả mã thoát `2`.

---

## 9. Kiểm thử offline

Không cần tài khoản GDT:

```powershell
# Windows
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1

# Linux / macOS bản ổn định
pwsh -NoProfile -File tests/Run-Tests.ps1

# macOS bản preview
pwsh-preview -NoProfile -File tests/Run-Tests.ps1
```

---

## Bảo mật

- `.env` đã được Git bỏ qua; không commit mật khẩu hoặc dữ liệu hóa đơn thật.
- Mật khẩu không được ghi vào log hoặc file Excel.
- Token đăng nhập chỉ được tạo tự động trong bộ nhớ khi chương trình chạy; người dùng không cần nhập token.

Chi tiết ở `SECURITY.md`.

---

## Đóng góp

| Tài liệu | Dành cho |
|---|---|
| [`AGENTS.md`](AGENTS.md) | AI coding agent: bản đồ repo, quy tắc kỹ thuật, Definition of Done |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | Quy trình, đặt tên branch, commit, Pull Request |
| [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) | Luồng dữ liệu và vai trò từng module |
| [`docs/EXCEL_FORMAT.md`](docs/EXCEL_FORMAT.md) | Hợp đồng workbook `.xlsx` và quy tắc đóng gói OPC |
| [`SECURITY.md`](SECURITY.md) | Dữ liệu không được commit |

```bash
git switch -c fix/ten-rang-nho
# sửa code và thêm test vào tests/Run-Tests.ps1
pwsh -NoProfile -File tests/Run-Tests.ps1   # phải in "All tests passed."
git commit -m "fix: mô tả vấn đề"
git push -u origin fix/ten-rang-nho
```

Logic gốc của bản Excel add-in: [dieutx/TaiHoaDonDienTu](https://github.com/dieutx/TaiHoaDonDienTu).
