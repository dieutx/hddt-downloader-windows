# HDDT Downloader

Công cụ PowerShell để tải **XML hóa đơn điện tử** từ GDT và xuất thành file Excel.

- Tải hóa đơn **mua vào** và **bán ra**.
- Tự động lấy CAPTCHA và đăng nhập bằng tài khoản GDT.
- Có thể dùng proxy/VPN khi IP hiện tại bị GDT hạn chế.
- Xuất Excel dạng dữ liệu, không chứa VBA hoặc sheet `MENU`.
- Có thể chạy bằng `.env` hoặc nhập cấu hình trực tiếp.

![Giao diện HDDT Downloader v1.1.1](docs/images/hddt-desktop-v1.1.1.png)

---
## 📺 Video hướng dẫn

👉 [Xem video hướng dẫn sử dụng](https://www.youtube.com/watch?v=nLWvosy-IM0)


## 1. Chạy nhanh trên Windows

### App có giao diện — Windows 10 / Windows 11

Bản v1.1.1 mặc định là **một file EXE duy nhất**. Nhấn đúp
`HDDT-Downloader-v1.1.1-windows.exe` để chạy; không cần tự giải nén, cài đặt,
PowerShell 7, Node.js hoặc Excel. App dùng Windows PowerShell 5.1 và .NET
Framework có sẵn trên Windows 10/11.

Giao diện compact gồm header nhỏ, form hóa đơn và điều khiển tải hai cột,
tiến độ và nhật ký luôn nhìn thấy. Không cuộn toàn trang ở các kích thước
mục tiêu 1160×768, 1366×768, 1440×900, 1920×1080. Nhật ký chính hiển thị
4–6 dòng gần nhất tùy chiều cao; nút con mắt hiện/ẩn mật khẩu để kiểm tra
trước khi tải. **Mở rộng nhật ký** xem toàn bộ trong cửa
sổ riêng. Proxy mở trong panel nhỏ; chuyển tab giữ dữ liệu đã nhập.

File Excel/XML/log mặc định lưu tại thư mục **HDDT-Downloader** trong
Documents; có thể chọn nơi lưu khác trên giao diện. Cấu hình tải/nhật ký
được lưu tại `%LOCALAPPDATA%\HDDT-Downloader\hddt-settings.json`, dùng chung
khi đổi tên, di chuyển hoặc chạy phiên bản EXE mới. Không lưu mật khẩu.
Mã chương trình được EXE tự nạp vào thư mục tạm riêng khi chạy và dọn khi
thoát bình thường; người dùng không cần thao tác với các file bên trong.

App kiểm tra GitHub Releases bằng tác vụ nền khi mở và qua nút **Cập nhật**.
Khi có bản mới: bấm **Tải bản mới**, chọn nơi lưu, rồi nhấn đúp EXE vừa tải.
Không cần giải nén. Nội dung thay đổi mở bằng **Chi tiết** hoặc chỉ báo
**Bản mới** trên header. Màn hình thấp dùng chỉ báo header để dành chỗ cho
nhật ký. Thông báo có thể đóng. App vẫn hỗ trợ ZIP portable của release cũ.
Lỗi kiểm tra cập nhật không chặn tải hóa đơn/chuyển XML offline.

Nguồn phiên bản là `app-version.json`, đồng bộ vào metadata EXE. Khi phát hành,
đính kèm file `HDDT-Downloader-vX.Y.Z-windows.exe` vào GitHub Release có tag
`vX.Y.Z` tương ứng. Checksum SHA-256 của asset được kiểm tra nếu API GitHub
cung cấp. Không gửi tài khoản/mật khẩu/token GDT đến GitHub.

Kiểm thử phần cập nhật và bản EXE (dữ liệu giả, không gọi GDT):

```powershell
powershell -NoProfile -File tests/Run-UpdateTests.ps1
powershell -NoProfile -File tests/Run-SingleFileTests.ps1 -ExePath <duong-dan-exe>
powershell -NoProfile -STA -File tests/Run-LayoutTests.ps1 -PreviewDirectory <thu-muc-anh>
```

Có thể chạy giao diện WPF bằng cách nhấn đúp **`run-gui.cmd`**, hoặc:

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Start-HddtGui.ps1
```

- Tab **Tải từ GDT**: nhập tài khoản, mật khẩu, khoảng ngày và chiều hóa đơn.
- Tab **XML có sẵn → Excel**: chọn thư mục XML để chuyển sang Excel offline.
- Chọn file Excel đầu ra, bấm **Bắt đầu**, theo dõi tiến độ và nhật ký.
- **Dừng an toàn** chờ request hiện tại và xuất phần dữ liệu đã xử lý. Đóng cửa
  sổ khi đang chạy cũng yêu cầu dừng an toàn trước khi thoát.
- **Nạp cấu hình .env** dùng cấu hình có sẵn, gồm các tùy chọn nâng cao; không
  sửa file `.env`. Mật khẩu nhập trên giao diện không ghi ra file tạm.
- **Cài đặt nâng cao** cho phép chỉnh số kết nối XML ban đầu/tối đa, giãn cách
  request, thời gian phục hồi sau 429, giãn cách API danh sách, retry, timeout,
  số hóa đơn mỗi trang, tự điều tiết danh sách và ghi nhật ký. Có nút khôi phục
  mặc định; tham số được kiểm tra trước khi áp dụng cho lần chạy tiếp theo.
- Bật **Ghi nhớ cho lần mở app sau** để lưu các tham số này vào
  `hddt-settings.json` trong thư mục dữ liệu của app. File chỉ chứa tham số tải/nhật ký. Nếu không bật,
  thay đổi chỉ dùng trong phiên hiện tại; muốn lưu thì thư mục app cần có quyền
  ghi. Nạp `.env` sau đó sẽ ưu tiên giá trị trong `.env` cho phiên hiện tại.
- Giao diện luôn từ chối ghi đè Excel đã có, kể cả khi `.env` đặt
  `OVERWRITE_OUTPUT=true`. Tên đầu ra mặc định có thời điểm chạy.

Giao diện dùng WPF/.NET Framework và Windows PowerShell 5.1 có sẵn trên Windows;
không cần PowerShell 7 hoặc thư viện ngoài. Không cần cài Excel để tạo `.xlsx`.
Giao diện chỉ dành cho Windows; các lệnh CLI bên dưới tiếp tục dùng được trên
Windows, macOS và Linux.

Để tạo app portable với launcher **`HDDT-Downloader.exe`**:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-WindowsApp.ps1
```

Script in đường dẫn một EXE trong thư mục build mới ở `work/desktop/`.
Chỉ cần sao chép file EXE đó để sử dụng hoặc chia sẻ. Build dùng danh sách
mã nguồn tường minh, không kèm `.env`, cấu hình cá nhân hoặc hóa đơn.

Nếu cần gói thư mục/ZIP kiểu cũ để phát triển hoặc hỗ trợ release trước:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Build-WindowsApp.ps1 -FolderBundle
```

Bản chạy từ source hoặc `-FolderBundle` tiếp tục dùng cấu hình/output cạnh
script như trước. Dữ liệu ở các bản portable cũ được giữ nguyên; có thể nạp
lại `.env` hoặc chọn các thư mục dữ liệu cũ trên giao diện.

Kiểm thử giao diện và luồng offline (không dùng tài khoản GDT):

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\tests\Run-DesktopTests.ps1
```

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
FROM_DATE=31/12/2025
TO_DATE=01/01/2026
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

Khi tải XML gặp HTTP 500, chương trình tự gọi `/api/query/invoices/detail`
bằng token của phiên đăng nhập hiện tại. Nếu lấy được chi tiết, dữ liệu được
ghi vào `TongHopHD_Mua`/`TongHopHD_Ban` và `ChiTietHD_Mua`/`ChiTietHD_Ban`.
Các trường API trả `null` được giữ trống hoặc giữ giá trị đã có từ danh sách;
không tự suy ra số lượng, đơn giá hay tạo XML thay thế. Hai sheet `_XML` chỉ
chứa dữ liệu đọc từ XML thật.

`BaoCao_LoiTaiHD` vẫn ghi nhận XML HTTP 500, nhưng kết quả cuối là
`Da lay du lieu tu API detail` khi đã bổ sung dữ liệu thành công. Nếu API detail
cũng lỗi, dòng tổng hợp từ danh sách vẫn được giữ cùng thông báo lỗi chi tiết.
Không cần nhập token/cookie hoặc bật thêm tùy chọn.

JSON được đọc từ byte UTF-8 gốc để giữ đúng tiếng Việt trên PowerShell 5.1;
workbook dùng font Calibri. URL mẫu chưa ghép MST hoặc URL không hợp lệ trong
`LinkTraCuu` được giữ dưới dạng chữ, chỉ URL HTTP/HTTPS hợp lệ mới có hyperlink.

Sheet chi tiết `ChiTietHD_Mua`/`ChiTietHD_Ban` theo bố cục v6.7.4 của bản gốc: cột 1 là `Mẫu số hóa đơn`, tiếp theo là thông tin chung rồi tới từng dòng hàng hóa (34 cột). Mẫu số và MST người bán/mua được ghi dạng Text để giữ số `0` ở đầu.

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
| `MAX_RETRIES` | Số lần thử lại lỗi mạng/5xx |
| `HTTP_TIMEOUT_SECONDS` | Thời gian chờ một request |
| `XML_CONCURRENCY` | Số kết nối tải XML chạy song song (1-10, mặc định 4) |
| `XML_MAX_CONCURRENCY` | Trần kết nối và số worker tải XML (mặc định 4) |
| `XML_REQUEST_INTERVAL_MS` | Giãn cách giữa hai request XML (mặc định 800 ms) |
| `XML_RECOVERY_STEP_SECONDS` | Khoảng im 429 mỗi bước phục hồi (mặc định 10 giây) |
| `BROWSER_USER_AGENT` | User-Agent gửi kèm; để trống dùng mặc định của trình duyệt |
| `LOG_HTTP_PROFILE` | In log nhóm header an toàn của request để chẩn đoán |

Xem toàn bộ tuỳ chọn trong `.env.example`.

Tải XML chạy theo pipeline: `XML_MAX_CONCURRENCY` worker cùng lúc, mỗi request
chờ `XML_REQUEST_INTERVAL_MS`. Log ghi rõ số worker được mở, mỗi worker nhận
hóa đơn nào (`Worker x/y nhận hóa đơn i/n: ...`) và mỗi dòng tiến độ cho biết
worker nào vừa xong cùng số worker đang chạy (`Worker x/y | đang chạy z/y`).
Bị HTTP 429 thì hệ thống nghỉ theo `Retry-After`, giảm một kết nối (tối đa một lần mỗi
30s để một hóa đơn retry không kéo tụt cả pool) và tăng khoảng cách; khi đã im
429 đủ lâu (`XML_RECOVERY_STEP_SECONDS`, mặc định 10s mỗi bước), khoảng cách
được giảm một nửa về mức `XML_REQUEST_INTERVAL_MS` rồi mới tăng lại kết nối,
không bao giờ vượt `XML_MAX_CONCURRENCY`. Chạy với `XML_CONCURRENCY=1` để giữ
hành vi tuần tự cũ.

---

## 8. Lỗi và tạm dừng

- XML HTTP 500 được thử lấy dữ liệu bằng API detail; HTTP 504 và lỗi parse vẫn được ghi vào `BaoCao_LoiTaiHD`. Chương trình tiếp tục xử lý phần còn lại.
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
