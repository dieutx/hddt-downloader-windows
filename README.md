# HDDT Downloader — Tải hóa đơn điện tử GDT bằng PowerShell

Tự động tải **XML hóa đơn điện tử** từ `hoadondientu.gdt.gov.vn` và tạo file **Excel `.xlsx`** gồm bảng tổng hợp, chi tiết và báo cáo lỗi. Không cần Microsoft Excel, không cần cài thư viện, không cần nhập CAPTCHA tay.

> 💡 **Chỉ cần 3 bước**: cài PowerShell → cấu hình `.env` → chạy `run.cmd`. Xem [Bắt đầu nhanh](#-bắt-đầu-nhanh-windows).
<img width="945" height="633" alt="image" src="https://github.com/user-attachments/assets/23668d0d-c5bc-48ab-8643-44898c4e61c1" />

---

## Đây là công cụ gì?

- **Tải hóa đơn mua vào và bán ra** từ cả hai nguồn `query` (hóa đơn điện tử thường) và `sco-query` (hóa đơn máy tính tiền).
- **Hai cách đăng nhập**: dán token từ trình duyệt, **hoặc** điền tài khoản + mật khẩu — chương trình tự đọc CAPTCHA và tự đăng nhập lại khi token hết hạn.
- **Chịu được GDT giới hạn tốc độ**: bị HTTP 429 thì tự chờ dài dần (tối đa 10 phút/lần, 12 lần) rồi tiếp tục, không phải chạy lại từ đầu.
- **Dừng an toàn bằng Ctrl+C**: dữ liệu đã tải vẫn được ghi ra Excel.
- Cùng logic với dự án Excel VBA [TaiHoaDonDienTu](https://github.com/dieutx/TaiHoaDonDienTu) (nguồn gốc ý tưởng do thành viên **ongke0711** chia sẻ trên diễn đàn [Giải Pháp Excel](https://www.giaiphapexcel.com/diendan/threads/t%E1%BA%A3i-h%C3%B3a-%C4%91%C6%A1n-%C4%91i%E1%BB%87n-t%E1%BB%AD-https-hoadondientu-gdt-gov-vn-excel-vba.171723/)).

### Bộ công cụ "Tải dữ liệu thuế điện tử" của cùng tác giả

Cùng giải bài toán tự động hóa khai thác hệ thống thuế GDT (CAPTCHA, tra cứu, tải về), khác hệ thống và nền tảng:

| Công cụ | Nền tảng | Hệ thống GDT | Dữ liệu tải về |
|---|---|---|---|
| **repo này** | PowerShell — Windows / Linux / macOS | `hoadondientu.gdt.gov.vn` (Hóa đơn điện tử) | XML hóa đơn mua vào/bán ra → Excel |
| [tai-ho-so-thue-gtgt](https://github.com/dieutx/tai-ho-so-thue-gtgt) | Python 3.10+ | `dichvucong.gdt.gov.vn` (Cổng Dịch vụ công) | ZIP tờ khai + metadata CSV/JSON |
| [TaiHoaDonDienTu](https://github.com/dieutx/TaiHoaDonDienTu) | Excel VBA | `hoadondientu.gdt.gov.vn` | XML hóa đơn → Excel (nguồn gốc logic) |

Dự án cộng đồng, không phải sản phẩm chính thức của cơ quan thuế. Đừng bao giờ commit token, mật khẩu hay dữ liệu hóa đơn thật.

---

## 🚀 Bắt đầu nhanh (Windows)

**Bước 1 — Tải code về máy** (mở Command Prompt):

```bat
git clone https://github.com/dieutx/hddt-downloader-windows.git
cd hddt-downloader-windows
copy .env.example .env
notepad .env
```

> Chưa có Git? Chạy `winget install --id Git.Git -e --source winget` rồi mở lại CMD.

**Bước 2 — Điền `.env`** tối thiểu 5 dòng sau (cách dễ nhất là dùng tài khoản + mật khẩu):

```dotenv
GDT_USERNAME=0123456789        ; mã số thuế (đăng nhập GDT)
GDT_PASSWORD=mat_khau          ; mật khẩu — để trống GDT_TOKEN
FROM_DATE=01/09/2026           ; từ ngày (dd/MM/yyyy)
TO_DATE=30/09/2026             ; đến ngày
INVOICE_DIRECTION=both         ; mua vào + bán ra
```

**Bước 3 — Nhấn đúp `run.cmd`** (hoặc gõ `run.cmd` trong CMD). Xong!

Kết quả nằm trong thư mục `output\`:

```text
output\
  HoaDonDienTu.xlsx        ← mở bằng Excel: tổng hợp, chi tiết, lỗi
  xml\purchase\*.xml       ← XML hóa đơn mua vào
  xml\sold\*.xml           ← XML hóa đơn bán ra
  logs\*.log               ← nhật ký chi tiết từng bước
```

---

## 🐧 Chạy trên Linux / macOS

Mã nguồn là PowerShell thuần nên chạy được trên Linux/macOS với **PowerShell 7 (pwsh)**. Không cần cài thêm thư viện nào khác — mọi thứ (giải nén ZIP, ghi Excel Open XML, gọi HTTPS) đều dùng sẵn trong .NET.

**Bước 1 — Cài PowerShell 7:**

```bash
# Ubuntu / Debian: cài trực tiếp từ bản phát hành của Microsoft
curl -sL https://github.com/PowerShell/PowerShell/releases/latest/download/powershell-7.4.6-linux-x64.tar.gz -o /tmp/pwsh.tgz
sudo mkdir -p /opt/pwsh && sudo tar zxf /tmp/pwsh.tgz -C /opt/pwsh
sudo chmod +x /opt/pwsh/pwsh && sudo ln -s /opt/pwsh/pwsh /usr/local/bin/pwsh
pwsh --version    # phải in ra PowerShell 7.4.6

# macOS (Homebrew):
brew install --cask powershell

# Hoặc theo hướng dẫn chính thức cho các distro khác:
# https://learn.microsoft.com/vi-vn/powershell/scripting/install/installing-powershell-on-linux
```

**Bước 2 — Tải code và cấu hình** (giống Windows, chỉ khác lệnh copy):

```bash
git clone https://github.com/dieutx/hddt-downloader-windows.git
cd hddt-downloader-windows
cp .env.example .env
nano .env        # điền như Bước 2 của Windows ở trên
```

**Bước 3 — Chạy** (thay `run.cmd` bằng lệnh pwsh):

```bash
pwsh -NoProfile -File Invoke-Hddt.ps1          # tải hóa đơn từ GDT
pwsh -NoProfile -File Parse-LocalXml.ps1       # chỉ parse XML có sẵn
```

**Lưu ý khi chạy trên Linux/macOS:**

- Thư mục đầu ra dùng dấu `/` (ví dụ `OUTPUT_DIR=/home/ban/output`) — hoặc để mặc định `output` là được.
- File Excel vẫn tạo được bình thường, mở bằng LibreOffice Calc hay Excel đều được.
- Cần pwsh 7 trở lên (pwsh 5.x chỉ có trên Windows và sẽ không chạy được một số hàm).
- Đã kiểm tra: toàn bộ test offline và chức năng parse XML chạy xanh trên Ubuntu + PowerShell 7.4.

---

## 🔑 Cách đăng nhập

### Cách 1: Tài khoản + mật khẩu (khuyên dùng — để trống `GDT_TOKEN`)

```dotenv
GDT_TOKEN=
GDT_USERNAME=0123456789
GDT_PASSWORD=mat_khau
```

Chương trình tự lấy CAPTCHA từ API, tự nhận diện mã (không OCR, không nhập tay) và đăng nhập. CAPTCHA sai thì tự lấy mã mới, thử lại tối đa 3 lần. Nếu giữa phiên bị `401/403` (token hết hạn), chương trình tự đăng nhập lại rồi chạy tiếp — không dừng phiên.

### Cách 2: Dán token từ trình duyệt (khi không muốn để mật khẩu trong `.env`)

1. Đăng nhập tại [hoadondientu.gdt.gov.vn](https://hoadondientu.gdt.gov.vn/).
2. Nhấn `F12` → tab **Console** → dán đoạn JavaScript sau rồi nhấn `Enter`:

```javascript
const t = (document.cookie.match(/eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/) || [])[0];
t ? console.log("Bearer " + t) : console.log("❌ Không tìm thấy token");
```

3. Sao chép chuỗi bắt đầu bằng `Bearer ` và điền vào `.env`:

```dotenv
GDT_TOKEN=Bearer eyJ...
GDT_USERNAME=
GDT_PASSWORD=
```

> ⚠️ Token là thông tin đăng nhập nhạy cảm: chỉ chạy đoạn mã trên đúng website GDT, không gửi token cho ai. Lỗi 401/403 với token dán tay thì lấy token mới (đăng nhập lại hoặc tải lại trang).

---

## 📋 Chỉnh gì trong `.env`?

### Tải từ GDT (dùng với `run.cmd`)

| Biến | Giá trị | Ý nghĩa |
|---|---|---|
| `GDT_TOKEN` / `GDT_USERNAME` + `GDT_PASSWORD` | | Chọn **một** trong hai cách đăng nhập |
| `FROM_DATE` / `TO_DATE` | `dd/MM/yyyy` | Khoảng ngày lập hóa đơn cần tải |
| `INVOICE_DIRECTION` | `purchase` / `sold` / `both` | Mua vào / bán ra / cả hai |
| `INCLUDE_REGULAR` | `true`/`false` | Lấy hóa đơn điện tử thường (nguồn `query`) |
| `INCLUDE_SCO` | `true`/`false` | Lấy hóa đơn máy tính tiền (nguồn `sco-query`) |
| `FETCH_RELATED` | `true`/`false` | Lấy thêm chuỗi hóa đơn thay thế/điều chỉnh (trạng thái 2-6) |
| `REDOWNLOAD_XML` | `true`/`false` | `false` = dùng lại XML đã tải, chỉ tải file còn thiếu |
| `PAGE_SIZE` | `50` | Số hóa đơn mỗi trang danh sách |
| `REQUEST_DELAY_MS` | `600` | Nghỉ giữa hai request — **đừng đặt thấp hơn** vì GDT sẽ 429 |
| `ADAPTIVE_THROTTLE` | `true` | Tự giãn cách khi bị 429, tự nới lỏng khi ổn định |
| `MAX_RETRIES` | `4` | Retry cho lỗi mạng/5xx (riêng 429 tự retry tới 12 lần) |
| `HTTP_TIMEOUT_SECONDS` | `90` | Chờ tối đa mỗi request |

### Đầu ra và log

| Biến | Giá trị | Ý nghĩa |
|---|---|---|
| `OUTPUT_DIR` | `output` | Thư mục chứa XML, Excel, log |
| `OUTPUT_XLSX` | `HoaDonDienTu.xlsx` | Tên file Excel |
| `OVERWRITE_OUTPUT` | `true` | Ghi đè file Excel cùng tên |
| `LOG_LEVEL` | `info` | `debug` / `info` / `warn` / `error` |
| `PROGRESS_EVERY` | `1` | In tiến độ mỗi N hóa đơn |

### Chỉ parse XML có sẵn (dùng với `parse-local.cmd`) — không cần token

| Biến | Giá trị | Ý nghĩa |
|---|---|---|
| `LOCAL_XML_DIR` | `output\xml` | Thư mục chứa XML (đọc cả thư mục con) |
| `LOCAL_DIRECTION` | `auto` | `auto` tự nhận diện theo thư mục/tên file; `purchase`/`sold` ép cả thư mục |
| `LOCAL_OUTPUT_XLSX` | `HoaDonDienTu_Local.xlsx` | Tên file Excel kết quả |

---

## ❓ Các tình huống thường gặp

**GDT trả 429 liên tục — phải làm gì?**
Không cần làm gì: chương trình tự chờ 15s → 30s → 60s → ... → 600s (tối đa 12 lần) rồi tiếp tục. Chỉ khi vượt 12 lần thử liên tiếp mới dừng, và dữ liệu đã tải vẫn được xuất ra Excel. Muốn ít bị 429 hơn thì tăng `REQUEST_DELAY_MS`.

**Tải bị mất điện / Ctrl+C nhầm — chạy lại có bị trùng không?**
Không: đặt `REDOWNLOAD_XML=false` rồi chạy lại — XML đã có sẽ được dùng lại, chỉ tải hóa đơn còn thiếu.

**Cần lại chuỗi hóa đơn bị thay thế/điều chỉnh?**
Giữ `FETCH_RELATED=true`. Kết quả nằm ở cột `RelatedChain`, `RelatedInfo`, `Original*` trong sheet `TongHop`. Nếu hóa đơn không ở trạng thái 2-6 thì không có gì để lấy, cột sẽ trống.

**Muốn tạm dừng?**
Ctrl+C lần 1: request hiện tại chạy xong rồi dừng, dữ liệu vẫn được ghi ra Excel. Ctrl+C lần 2: dừng ngay (không xuất Excel).

**Chỉ có file XML, muốn tạo Excel?**
Dùng chức năng 2: đặt `LOCAL_XML_DIR` trỏ tới thư mục XML rồi chạy `parse-local.cmd` (Linux: `pwsh -NoProfile -File Parse-LocalXml.ps1`). Không cần token, không gọi GDT.

**Chạy trên Linux/macOS được không?**
Được — cài [PowerShell 7](#-chạy-trên-linux--macos) rồi chạy bằng `pwsh -NoProfile -File Invoke-Hddt.ps1`.

**Báo lỗi 401/403?**
Token hết hạn hoặc sai tài khoản. Dùng cách đăng nhập tài khoản để chương trình tự đăng nhập lại; hoặc lấy token mới từ trình duyệt.

---

## 🧪 Kiểm thử offline (không cần token)

```powershell
# Windows
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1

# Linux / macOS
pwsh -NoProfile -File tests/Run-Tests.ps1
```

Test không gọi mạng và không cần token thật.

## Bảo mật

- Mật khẩu chỉ giữ trong bộ nhớ phiên chạy, không ghi vào log hay file Excel.
- Authorization header chỉ gửi tới host GDT cố định trong mã nguồn.
- Đừng commit `.env`, token hay dữ liệu hóa đơn thật vào Git (`.env` đã nằm trong `.gitignore`).
