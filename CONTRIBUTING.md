# Hướng dẫn đóng góp

Cảm ơn bạn muốn đóng góp cho `hddt-downloader-windows`. Công cụ đang chạy ổn
định và được xuất workbook data-only, nên Pull Request nên nhỏ, kiểm tra được và
chỉ thay đổi đúng phạm vi cần thiết.

Người đóng góp là người hoặc là AI Coding Agent đều tuân theo tài liệu này. Nếu
bạn là agent, hãy đọc thêm `AGENTS.md` (bản đồ repository, quy tắc kỹ thuật và
Definition of Done).

## Quy trình

```text
Fork repository
    ↓
Clone fork
    ↓
Tạo branch
    ↓
Sửa + chạy tests/Run-Tests.ps1
    ↓
Commit (một mục đích)
    ↓
Push
    ↓
Mở Pull Request kèm cách kiểm thử
```

```bash
git clone https://github.com/<tai-khoan>/hddt-downloader-windows.git
cd hddt-downloader-windows
git switch -c fix/ten-rang-nho
```

## Đặt tên branch

| Tiền tố | Dùng khi |
|---|---|
| `fix/` | sửa lỗi |
| `feat/` | tính năng mới đã thống nhất |
| `docs/` | tài liệu |
| `refactor/` | refactor nhỏ, giữ nguyên behavior |
| `test/` | kiểm thử hoặc fixture |
| `chore/` | bảo trì repository |

## Phạm vi thay đổi

- Đọc `README.md`, `AGENTS.md` và tài liệu trong `docs/` trước khi sửa.
- Không viết lại toàn bộ chương trình, không đổi endpoint hoặc luồng đăng nhập
  nếu Issue/PR không yêu cầu rõ ràng.
- Không thêm thư viện ngoài; chỉ dùng .NET tích hợp sẵn.
- Giữ tương thích **Windows PowerShell 5.1** lẫn **PowerShell 7** trên Windows,
  macOS và Linux.
- Giữ nguyên tên sheet, tên cột và thứ tự cột của workbook (xem
  `docs/EXCEL_FORMAT.md`); đây là hợp đồng với người dùng và với bản gốc
  `TaiHoaDonDienTu`.
- Ghi technical debt chưa xử lý vào mục "Đề xuất" của PR hoặc Issue riêng.

## Kiểm thử

```powershell
# PowerShell 7 (Windows, macOS, Linux)
pwsh -NoProfile -File tests/Run-Tests.ps1

# Windows PowerShell 5.1
powershell -NoProfile -File tests/Run-Tests.ps1
```

Đầu ra mong đợi: `All tests passed.`. Nếu không chạy được môi trường nào, ghi rõ
lệnh đã chạy, kết quả và lý do bỏ sót trong PR.

Ngoài test, nên kiểm tra thủ công khi sửa phần xuất Excel:

```powershell
# Xuất workbook từ XML có sẵn, không cần mạng
$env:LOCAL_XML_DIR = 'output/xml'
pwsh -NoProfile -File Parse-LocalXml.ps1
```

Rồi mở file `.xlsx` bằng Excel hoặc LibreOffice: 8 sheet, link tra cứu bấm được,
không có cảnh báo sửa chữa file.

## Commit convention

Ưu tiên Conventional Commits:

```text
fix: sửa lỗi đóng gói workbook thiếu quan hệ gốc
feat: thêm sheet LinkTraCuu vào workbook xuất
docs: mô tả luồng dữ liệu
test: thêm test cho tra cứu link
refactor: tách hàm dựng worksheet
chore: cập nhật .env.example
```

Một commit có một mục đích rõ ràng; không trộn thay đổi không liên quan.

## Pull Request

PR nên:

- Tập trung vào một vấn đề chính.
- Mô tả hiện tượng, nguyên nhân, cách sửa và phạm vi ảnh hưởng.
- Ghi lệnh kiểm thử đã chạy và kết quả.
- Nêu rõ thay đổi behavior (đặc biệt khi đổi tên cột, thứ tự cột, nhãn
  `tthai`/`ttxly`, câu chữ link tra cứu).
- Kèm test hồi quy cho mỗi bug đã sửa.

Dùng mẫu `.github/PULL_REQUEST_TEMPLATE.md`.

## Dữ liệu tuyệt đối không được commit

- `.env`, username, password, token, cookie, `Authorization` header.
- XML/PDF/ZIP hóa đơn thật, MST thật của tài khoản đăng nhập.
- Tên, địa chỉ, số điện thoại hoặc thông tin khách hàng.
- Workbook `.xlsx` đã xuất và `output/`.

Fixture trong `tests/fixtures/` phải là dữ liệu giả. MST trong bảng
`src/LinkTraCuu.ps1` là dữ liệu tham chiếu công khai (nhà cung cấp T-VAN) nên
được giữ. Chi tiết ở `SECURITY.md`.
