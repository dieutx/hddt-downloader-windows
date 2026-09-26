## Vấn đề

<!-- Hiện tượng gặp phải, kèm lỗi nguyên văn nếu có. -->

## Nguyên nhân

<!-- Vì sao xảy ra. Nếu chưa rõ, ghi "chưa xác định" thay vì đoán. -->

## Cách sửa

<!-- Những gì đã đổi, ngắn gọn. -->

## Phạm vi ảnh hưởng

<!-- Đổi behavior, tên cột, thứ tự cột, câu chữ, hay chỉ sửa lỗi? -->

## Cách kiểm thử

```powershell
pwsh -NoProfile -File tests/Run-Tests.ps1
```

<!-- Kết quả thực tế đã chạy, và kiểm tra thủ công nếu có (ví dụ mở file .xlsx). -->

## Checklist

- [ ] Không có credential, token, XML/ZIP hóa đơn thật hay MST thật trong diff
- [ ] Không sửa tên sheet, tên cột, thứ tự cột ngoài phạm vi task (xem `docs/EXCEL_FORMAT.md`)
- [ ] Thêm test hồi quy cho bug đã sửa
- [ ] Cập nhật `README.md`/`docs/` nếu behavior đổi
- [ ] Tương thích Windows PowerShell 5.1 và PowerShell 7

## Đề xuất (ngoài phạm vi task)

<!-- Vấn đề phát hiện thêm nhưng không sửa trong PR này. -->
