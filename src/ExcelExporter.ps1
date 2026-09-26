Set-StrictMode -Version 2.0

# Data-only Excel export.  The VBA project and the MENU/reference worksheets
# from TaiHoaDonDienTu are deliberately not copied into the generated workbook.
. (Join-Path $PSScriptRoot 'LinkTraCuu.ps1')

$script:SpreadsheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:RelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$script:PackageRelationshipNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships'
$script:ContentTypeNamespace = 'http://schemas.openxmlformats.org/package/2006/content-types'

$script:SourceSummaryHeaders = @(
    'STT', 'Tên HĐ', 'Mẫu HĐ', 'Ký hiệu hóa đơn', 'Số hóa đơn',
    'Ngày lập hóa đơn', 'Đơn vị tiền tệ', 'Tỷ giá', 'Tên người bán',
    'MST người bán', 'Địa chỉ người bán', 'Ngày ký số người bán', 'Mã CQT',
    'Ngày cấp mã CQT', 'Tên người mua', 'MST người mua', 'Địa chỉ người mua',
    'Thuế suất', 'Thành tiền`ntrước thuế', 'Tiền thuế', 'Giảm trừ thuế suất',
    'Thuế suất 2', 'Thành tiền `ntrước thuế 2', 'Tiền thuế 2', 'Giảm trừ thuế suất 2',
    'Thuế suất 3', 'Thành tiền `ntrước thuế 3', 'Tiền thuế 3', 'Giảm trừ thuế suất 3',
    'Thuế suất 4', 'Thành tiền`ntrước thuế 4', 'Tiền thuế 4', 'Giảm trừ thuế suất 4',
    'Thuế suất 5', 'Thành tiền `ntrước thuế 5', 'Tiền thuế 5', 'Giảm trừ thuế suất 5',
    'Thuế suất 6', 'Thành tiền `ntrước thuế 6', 'Tiền thuế 6', 'Giảm trừ thuế suất 6',
    'Tổng tiền chưa thuế`n(Tổng cộng thành tiền chưa có thuế)', 'Tổng giảm trừ không chịu thuế',
    'Tổng tiền thuế (Tổng cộng tiền thuế)', 'Tên loại phí', 'Tổng tiền phí',
    'Tổng tiền chiết khấu thương mại', 'Tổng giảm trừ khác',
    'Tổng tiền thanh toán bằng số', 'Tổng tiền thanh toán bằng chữ', 'Ghi chú',
    'Trạng thái hóa đơn', 'Kết quả kiểm tra hóa đơn', 'MSTTCGP', 'Link tra cứu', 'Mã tra cứu',
    'Chuỗi hóa đơn liên quan', 'Loại hóa đơn gốc', 'Ký hiệu mẫu số HĐ gốc',
    'Ký hiệu HĐ gốc', 'Số HĐ gốc', 'Ngày lập HĐ gốc', 'Ghi chú HĐ gốc', 'Thông tin liên quan'
)

$script:SourceDetailHeaders = @(
    'Ký hiệu HĐ', 'Số hóa đơn', 'Ngày lập hóa đơn', 'Đơn vị tiền tệ', 'Tỷ giá',
    'Tên người bán', 'MST người bán', 'Địa chỉ người bán', 'Ngày ký số người bán',
    'Mã CQT', 'Ngày cấp mã CQT', 'Tên người mua', 'MST người mua', 'Địa chỉ người mua',
    'Số thứ tự', 'Tính chất', 'Mã HHDV', 'Tên HHDV', 'DVT', 'Số lượng', 'Đơn giá',
    'Tỷ lệ chiết khấu', 'Số tiền chiết khấu', 'Loại thuế suất',
    'Thuế suất`n(TS thuế VAT)', 'Thành tiền`n(Chưa thuế VAT)', 'Tiền thuế`n(Tiền thuế GTGT)',
    'Thành tiền có thuế (Thành tiền có thuế GTGT)', 'Tiền thuế`ntrên HĐ', 'Chênh lệch kê khai thuế',
    'MSTTCGP', 'Link tra cứu', 'Mã tra cứu'
)
for ($headerIndex = 0; $headerIndex -lt $script:SourceSummaryHeaders.Count; $headerIndex++) {
    $script:SourceSummaryHeaders[$headerIndex] = $script:SourceSummaryHeaders[$headerIndex].Replace('`n', "`n")
}
for ($headerIndex = 0; $headerIndex -lt $script:SourceDetailHeaders.Count; $headerIndex++) {
    $script:SourceDetailHeaders[$headerIndex] = $script:SourceDetailHeaders[$headerIndex].Replace('`n', "`n")
}

$script:SourceXmlHeadersPurchase = @(
    'Id', 'Ký hiệu hóa  đơn', 'Số hóa đơn', 'Ngày lập hóa đơn', 'Đơn vị tiền tệ', 'Tỷ giá',
    'Tên người bán', 'MST người bán', 'Địa chỉ người bán', 'Ngày ký số người bán',
    'Mã CQT', 'Ngày cấp mã CQT', 'Tên người mua', 'MST người mua', 'Địa chỉ người mua',
    'Số thứ tự', 'Mã HHDV', 'Tên HHDV', 'DVT', 'Số lượng', 'Đơn giá', 'Tỷ lệ chiết khấu',
    'Số tiền chiết khấu', 'Thành tiền`n(Chưa thuế VAT)', 'Thuế suất`n(TS thuế VAT)',
    'Tiền thuế`n(Tiền thuế GTGT)', 'Thành tiền có thuế (Thành tiền có thuế GTGT)',
    'Tiền thuế`ntrên HĐ', 'MSTTCGP', 'Link tra cứu', 'Mã tra cứu', 'Link tải hóa đơn gốc'
)
$script:SourceXmlHeadersSold = @(
    'Id', 'Ký hiệu hóa  đơn', 'Số hóa đơn', 'Ngày lập hóa đơn', 'Đơn vị tiền tệ', 'Tỷ giá',
    'Tên người mua', 'MST người mua', 'Địa chỉ người mua', 'Ngày ký số người mua',
    'Mã CQT', 'Ngày cấp mã CQT', 'Tên người bán', 'MST người bán', 'Địa chỉ người bán',
    'Số thứ tự', 'Mã HHDV', 'Tên HHDV', 'DVT', 'Số lượng', 'Đơn giá', 'Tỷ lệ chiết khấu',
    'Số tiền chiết khấu', 'Thành tiền`n(Chưa thuế VAT)', 'Thuế suất`n(TS thuế VAT)',
    'Tiền thuế`n(Tiền thuế GTGT)', 'Thành tiền có thuế (Thành tiền có thuế GTGT)',
    'Tiền thuế`ntrên HĐ', 'MSTTCGP', 'Link tra cứu', 'Mã tra cứu', 'Link tải hóa đơn gốc'
)
for ($headerIndex = 0; $headerIndex -lt $script:SourceXmlHeadersPurchase.Count; $headerIndex++) {
    $script:SourceXmlHeadersPurchase[$headerIndex] = $script:SourceXmlHeadersPurchase[$headerIndex].Replace('`n', "`n")
    $script:SourceXmlHeadersSold[$headerIndex] = $script:SourceXmlHeadersSold[$headerIndex].Replace('`n', "`n")
}
$script:SourceXmlWidths = @(
    11.85546875, 9.140625, 12.28515625, 15.85546875, 9.140625, 8, 49.5703125,
    19.140625, 30.42578125, 16.28515625, 40.85546875, 27.5703125, 37.7109375,
    15.28515625, 19.42578125, 9.140625, 13.85546875, 33.28515625, 14.85546875,
    13.7109375, 18.42578125, 9.140625, 14.7109375, 22.85546875, 15.140625,
    18.7109375, 22.85546875, 16, 20.140625, 40.85546875, 27.42578125,
    36.42578125
)

$script:SourceErrorHeaders = @(
    'STT', 'Thời gian ghi nhận', 'Loại hóa đơn', 'Nguồn API', 'Mã số thuế người bán',
    'Ký hiệu mẫu số', 'Ký hiệu hóa đơn', 'Số hóa đơn', 'Ngày lập hóa đơn',
    'Công đoạn lỗi', 'Endpoint', 'HTTP status', 'Nội dung lỗi', 'Số lần đã thử',
    'Retry-After', 'Kết quả cuối', 'Ghi chú'
)

$script:SourceSummaryWidths = @(
    8.7109375, 31.42578125, 9.85546875, 11.140625, 12, 13.85546875, 9.28515625,
    8.140625, 51.5703125, 19.85546875, 40.85578125, 15.42578125, 47.42578125,
    13.28515625, 37.5703125, 18.28515625, 38.28515625, 10.7109375, 16.28515625,
    16.28515625, 16.28515625, 12.42578125, 17.5703125, 17.5703125, 17.5703125,
    10.7109375, 19.42578125, 17.5703125, 17.5703125, 10.7109375, 19.42578125,
    17.5703125, 17.140625, 10.7109375, 15.28515625, 13.28515625, 15.28515625,
    10.7109375, 15.28515625, 15.28515625, 15.28515625, 25, 19.42578125, 19.42578125,
    19.42578125, 53.42578125, 19.42578125, 19.42578125, 19.42578125, 19.42578125,
    16, 26.5703125, 69.7109375, 19.140625, 54.42578125, 35.42578125, 32.7109375,
    18.7109375, 18.7109375, 18.7109375, 18.7109375, 18.7109375, 18.7109375, 45.7109375
)
$script:SourceDetailWidths = @(
    9.140625, 12.28515625, 15.85546875, 9.140625, 8, 49.5703125, 19.140625,
    30.42578125, 16.28515625, 40.85546875, 27.5703125, 37.7109375, 15.28515625,
    19.42578125, 9.140625, 9.140625, 11.5703125, 33.28515625, 14.85546875,
    13.7109375, 13.7109375, 9.140625, 14.7109375, 14.7109375, 11.5703125,
    22.85546875, 18.7109375, 22.85546875, 16, 19.42578125, 16.85546875,
    47.28515625, 29.7109375
)
$script:SourceErrorWidths = @(
    7.7109375, 19.7109375, 14.7109375, 12.7109375, 20.7109375, 18.7109375,
    18.7109375, 16.7109375, 16.7109375, 18.7109375, 45.7109375, 12.7109375,
    45.7109375, 14.7109375, 13.7109375, 24.7109375, 35.7109375
)

# Nhãn trạng thái hóa đơn và kết quả kiểm tra lấy từ sheet LinkTraCuu
# (I2:I8 và O2:O11) nên không thể lệch khỏi bảng tra cứu.
$script:SourceStatusLabels = @(Get-TaiHoaDonDienTuStatusLabels)
$script:SourceValidationLabels = @(Get-TaiHoaDonDienTuValidationLabels)

# Sheet LinkTraCuu: bảng định tuyến link tra cứu được ghi kèm workbook để
# người dùng tự thêm/sửa, và chính bảng này sinh ra link ở cột 55/56.
$script:SourceLookupHeaders = @(Get-TaiHoaDonDienTuLookupHeaders)
$script:SourceLookupWidths = @(
    46, 14, 14, 62, 22, 40, 3, 10, 30, 3, 10, 46, 3, 10, 60
)

function Get-ExcelColumnName {
    param([Parameter(Mandatory = $true)][int]$Index)
    if ($Index -lt 1) { throw 'Column index must be at least 1.' }
    $name = ''
    while ($Index -gt 0) {
        $Index--
        $name = [char](65 + ($Index % 26)) + $name
        $Index = [Math]::Floor($Index / 26)
    }
    return $name
}

function New-Utf8XmlWriter {
    param([Parameter(Mandatory = $true)][string]$Path)
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.OmitXmlDeclaration = $false
    return [System.Xml.XmlWriter]::Create($Path, $settings)
}

function Write-XmlStartElement {
    param($Writer, [string]$Name, [string]$Namespace = $script:SpreadsheetNamespace)
    [void]$Writer.WriteStartElement($Name, $Namespace)
}

function Write-TextFileUtf8NoBom {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Content)
    [IO.File]::WriteAllText($Path, $Content, (New-Object Text.UTF8Encoding($false)))
}

function Get-ExcelObjectValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary] -and $Object.Contains($Name)) {
        if ($null -eq $Object[$Name]) { return $Default }
        return $Object[$Name]
    }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-ExcelText {
    param($Object, [string]$Name, [string]$Default = '')
    $value = Get-ExcelObjectValue $Object $Name $Default
    if ($null -eq $value) { return $Default }
    if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) { return $Default }
    return [string]$value
}

function ConvertTo-ExcelDate {
    param($Value)
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return $null }
    if ($Value -is [datetime]) { return ([datetime]$Value).Date }
    $text = ([string]$Value).Trim()
    $date = [datetime]::MinValue
    $styles = [Globalization.DateTimeStyles]::AllowWhiteSpaces
    if ($text -match '^\d{1,2}/\d{1,2}/\d{4}$') {
        if ([datetime]::TryParse($text, [Globalization.CultureInfo]::GetCultureInfo('vi-VN'), $styles, [ref]$date)) { return $date.Date }
    }
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$date)) { return $date.Date }
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::GetCultureInfo('vi-VN'), $styles, [ref]$date)) { return $date.Date }
    return $null
}

function ConvertTo-ExcelNumber {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        return [Convert]::ToDouble($Value, [Globalization.CultureInfo]::InvariantCulture)
    }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $number = 0.0
    $styles = [Globalization.NumberStyles]::AllowThousands -bor [Globalization.NumberStyles]::AllowDecimalPoint
    $vietnameseDecimal = ($text -match ',\d{1,2}$' -and $text -notmatch '\.') -or ($text -match ',\d{1,2}$' -and $text -match '\.')
    if ($vietnameseDecimal) {
        if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('vi-VN'), [ref]$number)) { return $number }
    }
    if ([double]::TryParse($text, $styles, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) { return $number }
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('vi-VN'), [ref]$number)) { return $number }
    return $Value
}

function ConvertTo-ExcelTaxRate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ($text -in @('KKKNT', 'KCT')) { return $text }
    $isPercent = $text.EndsWith('%')
    $numberText = if ($isPercent) { $text.Substring(0, $text.Length - 1).Trim() } else { $text }
    $number = ConvertTo-ExcelNumber $numberText
    if ($number -is [string]) { return $Value }
    if ($isPercent) { return ([double]$number) / 100.0 }
    if ([double]$number -gt 1) { return ([double]$number) / 100.0 }
    return [double]$number
}

function Get-ExcelRawOrFallback {
    param($Raw, $Fallback, [string]$Name, [string[]]$FallbackNames = @())
    $value = Get-ExcelObjectValue $Raw $Name $null
    $hasValue = $null -ne $value
    if ($hasValue -and $value -is [string]) { $hasValue = -not [string]::IsNullOrWhiteSpace($value) }
    if ($hasValue -and $value -is [System.Collections.IEnumerable] -and $value -isnot [string]) { $hasValue = @($value).Count -gt 0 }
    if ($hasValue) { return $value }
    foreach ($fallbackName in $FallbackNames) {
        $value = Get-ExcelObjectValue $Fallback $fallbackName $null
        $hasValue = $null -ne $value
        if ($hasValue -and $value -is [string]) { $hasValue = -not [string]::IsNullOrWhiteSpace($value) }
        if ($hasValue -and $value -is [System.Collections.IEnumerable] -and $value -isnot [string]) { $hasValue = @($value).Count -gt 0 }
        if ($hasValue) { return $value }
    }
    return $null
}

function Get-ExcelAdditionalItems {
    param($Value)
    if ($null -eq $Value) { return @() }
    return @($Value)
}

function Get-ExcelLookupCode {
    param($Containers)
    foreach ($container in (Get-ExcelAdditionalItems $Containers)) {
        foreach ($item in (Get-ExcelAdditionalItems $container)) {
            $field = Get-ExcelObjectValue $item 'TTruong' (Get-ExcelObjectValue $item 'Field' '')
            if (Test-TaiHoaDonDienTuLookupField ([string]$field)) {
                $value = Get-ExcelObjectValue $item 'DLieu' (Get-ExcelObjectValue $item 'Value' '')
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
            }
        }
    }
    return ''
}

# Chọn MST dùng để tra bảng LinkTraCuu.  Ưu tiên MSTTCGP (nhà cung cấp T-VAN
# do GDT trả về).  Khi GDT không trả MSTTCGP - hóa đơn phát hành trực tiếp
# qua MSSVĐHĐN, hoặc XML không có TTChung/MSTTCGP - thử lần lượt MST của các
# bên liên quan; nếu bên đó đúng là một nhà cung cấp trong bảng LinkTraCuu thì
# trang tra cứu của họ vẫn dùng được cho hóa đơn này.
function Resolve-ExcelLookupTaxCode {
    param([string]$ProviderTaxCode, [string[]]$CounterPartyTaxCodes)
    if (-not [string]::IsNullOrWhiteSpace($ProviderTaxCode)) {
        return $ProviderTaxCode.Trim()
    }
    foreach ($taxCode in @(Get-ExcelAdditionalItems $CounterPartyTaxCodes)) {
        $candidate = ([string]$taxCode).Trim()
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if (-not [string]::IsNullOrWhiteSpace((Get-TaiHoaDonDienTuSellerLookupLink $candidate))) { return $candidate }
        if (-not [string]::IsNullOrWhiteSpace((Get-TaiHoaDonDienTuLookupLink $candidate))) { return $candidate }
    }
    return ''
}

function Get-ExcelLookupLink {
    param(
        [string]$ProviderTaxCode,
        [string]$SellerTaxCode,
        [string]$TaxAuthorityCode,
        [string]$InvoiceId,
        [string[]]$CounterPartyTaxCodes = @(),
        [ValidateSet('summary', 'detail', 'xml')][string]$Mode = 'summary'
    )
    $lookupTaxCode = Resolve-ExcelLookupTaxCode -ProviderTaxCode $ProviderTaxCode -CounterPartyTaxCodes $CounterPartyTaxCodes
    if ([string]::IsNullOrWhiteSpace($lookupTaxCode)) {
        if ($SellerTaxCode -in @('0110269067', '0110269067-002')) {
            return 'https://gsm-einvoice.hilo.com.vn/'
        }
        if ($Mode -eq 'xml') { return 'Khong tim thay link tra cuu' }
        return 'Khong co link tra cuu'
    }

    # Đường dẫn tìm được nhờ MST bên liên quan: dùng link của chính MST đó,
    # kể cả link riêng theo người bán (cột C của sheet LinkTraCuu).
    if ([string]::IsNullOrWhiteSpace($ProviderTaxCode)) {
        $sellerLink = Get-TaiHoaDonDienTuSellerLookupLink $lookupTaxCode
        if (-not [string]::IsNullOrWhiteSpace($sellerLink)) { return $sellerLink }
        $counterPartyLink = Get-TaiHoaDonDienTuLookupLink $lookupTaxCode
        if (-not [string]::IsNullOrWhiteSpace($counterPartyLink)) { return $counterPartyLink }
        if ($Mode -eq 'xml') { return 'Khong tim thay link tra cuu' }
        return 'Khong co link tra cuu'
    }

    switch ($ProviderTaxCode) {
        '0100684378' {
            if (-not [string]::IsNullOrWhiteSpace($TaxAuthorityCode)) {
                return 'https://{0}-tt78.vnpt-invoice.com.vn/?strFkey={1}' -f $SellerTaxCode, $TaxAuthorityCode
            }
            if ($Mode -eq 'xml') { return 'Khong co MCCQT' }
            return 'Khong co link tra cuu'
        }
        '0101360697' {
            if (-not [string]::IsNullOrWhiteSpace($InvoiceId)) {
                if ($Mode -eq 'summary' -and [string]::IsNullOrWhiteSpace($TaxAuthorityCode)) { return '' }
                return 'https://van.ehoadon.vn/Lookup?InvoiceGUID={0}' -f $InvoiceId
            }
            if ($Mode -eq 'summary') { return '' }
            return 'https://van.ehoadon.vn/Lookup?InvoiceGUID= [DLHDon Id]'
        }
        '0105987432' {
            return 'https://{0}hd.easyinvoice.com.vn' -f $SellerTaxCode
        }
        default {
            $link = Get-TaiHoaDonDienTuLookupLink $ProviderTaxCode
            if (-not [string]::IsNullOrWhiteSpace($link)) { return $link }
            return ''
        }
    }
}

function Get-ExcelProviderLookupCode {
    param(
        [string]$ProviderTaxCode,
        [string]$SellerTaxCode,
        [string]$TaxAuthorityCode,
        [string]$InvoiceId
    )
    switch ($ProviderTaxCode) {
        '0100684378' { return $TaxAuthorityCode }
        '0101360697' { return $InvoiceId }
        '0105987432' { return $SellerTaxCode }
        default { return '' }
    }
}

function Get-ExcelStatusLabel {
    param($Value)
    $number = 0
    if (-not [int]::TryParse([string]$Value, [ref]$number)) { return '' }
    if ($number -lt 0 -or $number -ge $script:SourceStatusLabels.Count) { return '' }
    return $script:SourceStatusLabels[$number]
}

function Get-ExcelValidationLabel {
    # Ô đầu tiên của O2:O11 là 'Tất cả', VBA tra bằng ttxly + 2 (mảng 1-based)
    # nên ttxly N ứng với chỉ số N + 1.
    param($Value)
    $number = 0
    if (-not [int]::TryParse([string]$Value, [ref]$number)) { return '' }
    $index = $number + 1
    if ($index -lt 0 -or $index -ge $script:SourceValidationLabels.Count) { return '' }
    return $script:SourceValidationLabels[$index]
}

function New-ExcelDataRow {
    return [pscustomobject]@{ Cells = @{}; Height = $null }
}

function Set-ExcelDataCell {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][int]$Column,
        $Value,
        [int]$Style = 5,
        [string]$Hyperlink = ''
    )
    $Row.Cells[$Column] = [pscustomobject]@{ Value = $Value; Style = $Style; Hyperlink = $Hyperlink }
}

function Get-ExcelRowHeight {
    param([string[]]$Texts)
    $lineCount = 1
    foreach ($text in $Texts) {
        if ($null -eq $text) { continue }
        $count = ([regex]::Matches([string]$text, "`n")).Count + 1
        if ($count -gt $lineCount) { $lineCount = $count }
    }
    $height = [Math]::Min(150.0, [Math]::Max(15.0, 15.0 * $lineCount))
    return $height
}

function Get-ExcelDetailGroupKey {
    param($Detail, [int]$Index)
    $summary = Get-ExcelObjectValue $Detail 'InvoiceSummary' $null
    $xmlFile = Get-ExcelText $summary 'XmlFile' (Get-ExcelText $Detail 'XmlFile' '')
    $invoiceId = Get-ExcelText $summary 'InvoiceId' (Get-ExcelText $Detail 'InvoiceId' '')
    $series = Get-ExcelText $summary 'InvoiceSeries' (Get-ExcelText $Detail 'InvoiceSeries' '')
    $number = Get-ExcelText $summary 'InvoiceNumber' (Get-ExcelText $Detail 'InvoiceNumber' '')
    $direction = Get-ExcelText $summary 'Direction' (Get-ExcelText $Detail 'Direction' '')
    $key = '{0}|{1}|{2}|{3}|{4}' -f $direction, $xmlFile, $invoiceId, $series, $number
    if ([string]::IsNullOrWhiteSpace($key.Replace('|', ''))) { return [string]$Index }
    return $key
}

function Get-ExcelDetailGroups {
    param($Rows)
    $groups = [ordered]@{}
    $index = 0
    foreach ($detail in @($Rows)) {
        if ($null -eq $detail) { continue }
        $key = Get-ExcelDetailGroupKey $detail $index
        $index++
        if (-not $groups.Contains($key)) {
            $groups[$key] = New-Object System.Collections.Generic.List[object]
        }
        $groups[$key].Add($detail)
    }
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $groups.GetEnumerator()) {
        $first = $entry.Value[0]
        $summary = Get-ExcelObjectValue $first 'InvoiceSummary' $null
        if ($null -eq $summary) { $summary = $first }
        $result.Add([pscustomobject]@{ Key = $entry.Key; Summary = $summary; Details = $entry.Value.ToArray() })
    }
    return $result.ToArray()
}

function New-ExcelSummaryRows {
    param($Rows)
    $result = New-Object System.Collections.Generic.List[object]
    $sequence = @{}
    foreach ($summary in @($Rows)) {
        if ($null -eq $summary) { continue }
        $direction = (Get-ExcelText $summary 'Direction' 'purchase').ToLowerInvariant()
        if ($direction -notin @('purchase', 'sold')) { $direction = 'purchase' }
        if (-not $sequence.ContainsKey($direction)) { $sequence[$direction] = 1 }
        $index = [int]$sequence[$direction]
        $sequence[$direction] = $index + 1

        $raw = Get-ExcelObjectValue $summary 'GdtIndex' $null
        $row = New-ExcelDataRow
        Set-ExcelDataCell $row 1 $index 5
        $map = @(
            @('tlhdon', @('InvoiceType')),
            @('khmshdon', @('TemplateCode')),
            @('khhdon', @('InvoiceSeries')),
            @('shdon', @('InvoiceNumber')),
            @('tdlap', @('InvoiceDate')),
            @('dvtte', @('Currency')),
            @('tgia', @('ExchangeRate')),
            @('nbten', @('SellerName')),
            @('nbmst', @('SellerTaxCode')),
            @('nbdchi', @('SellerAddress')),
            @('nky', @('SellerSigningTime')),
            @('mhdon', @('TaxAuthorityCode')),
            @('ncma', @('TaxAuthoritySigningTime')),
            @('nmten', @('BuyerName')),
            @('nmmst', @('BuyerTaxCode')),
            @('nmdchi', @('BuyerAddress'))
        )
        $targetColumns = 2..17
        for ($mapIndex = 0; $mapIndex -lt $map.Count; $mapIndex++) {
            $value = Get-ExcelRawOrFallback $raw $summary $map[$mapIndex][0] $map[$mapIndex][1]
            $column = $targetColumns[$mapIndex]
            $style = 5
            if ($column -in @(6, 12, 14)) {
                $value = ConvertTo-ExcelDate $value
                $style = 6
            }
            elseif ($column -in @(8, 19, 20, 21, 23, 24, 25, 27, 28, 29, 31, 32, 33, 35, 36, 37, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50)) {
                $value = ConvertTo-ExcelNumber $value
                $style = 7
            }
            elseif ($column -in @(3, 4, 5, 10, 16, 18, 22, 26, 30, 34, 38, 51, 54)) {
                $style = 16
            }
            Set-ExcelDataCell $row $column $value $style
        }

        $taxBuckets = Get-ExcelObjectValue $raw 'thttltsuat' @()
        for ($bucketIndex = 0; $bucketIndex -lt 6; $bucketIndex++) {
            $bucket = $null
            $buckets = @($taxBuckets)
            if ($bucketIndex -lt $buckets.Count) { $bucket = $buckets[$bucketIndex] }
            if ($null -eq $bucket) { continue }
            $baseColumn = 18 + (4 * $bucketIndex)
            Set-ExcelDataCell $row $baseColumn (ConvertTo-ExcelTaxRate (Get-ExcelText $bucket 'tsuat' '')) 8
            Set-ExcelDataCell $row ($baseColumn + 1) (ConvertTo-ExcelNumber (Get-ExcelText $bucket 'thtien' '')) 7
            Set-ExcelDataCell $row ($baseColumn + 2) (ConvertTo-ExcelNumber (Get-ExcelText $bucket 'tthue' '')) 7
            Set-ExcelDataCell $row ($baseColumn + 3) (Get-ExcelObjectNumber $bucket 'gttsuat') 7
        }

        $fee = @($(Get-ExcelObjectValue $raw 'thttlphi' @())) | Select-Object -First 1
        if ($null -ne $fee) {
            Set-ExcelDataCell $row 45 (Get-ExcelText $fee 'tlphi' '') 5
            Set-ExcelDataCell $row 46 (ConvertTo-ExcelNumber (Get-ExcelText $fee 'tphi' '')) 7
        }

        # Keep the source column positions verbatim.  In particular, the
        # template has separate tgtkhac (AV), tgtttbso (AW), tgtttbchu (AX)
        # and gchu (AY) columns; do not compact them.
        $totalFields = @(
            @('tgtcthue', 42), @('tgtkcthue', 43), @('tgtthue', 44), @('ttcktmai', 47),
            @('tgtkhac', 48), @('tgtttbso', 49), @('tgtttbchu', 50), @('gchu', 51),
            @('msttcgp', 54)
        )
        $totalFallbacks = @{
            tgtcthue = @('AmountBeforeTax')
            tgtkcthue = @('TaxExempt')
            tgtthue = @('TaxAmount')
            ttcktmai = @('CommercialDiscount')
            tgtkhac = @('OtherReduction')
            tgtttbso = @('TotalAmount')
            tgtttbchu = @('TotalAmountText')
            gchu = @('Note')
            msttcgp = @('ProviderTaxCode')
        }
        foreach ($field in $totalFields) {
            $fallbackNames = @($totalFallbacks[$field[0]])
            $value = Get-ExcelRawOrFallback $raw $summary $field[0] $fallbackNames
            $style = if ($field[1] -eq 51) { 5 } elseif ($field[1] -eq 54) { 16 } else { 7 }
            if ($field[0] -eq 'gchu') { $value = Get-ExcelText $raw 'gchu' (Get-ExcelText $summary 'Note' '') }
            Set-ExcelDataCell $row ([int]$field[1]) $value $style
        }

        $status = Get-ExcelRawOrFallback $raw $summary 'tthai' @('Status')
        $validation = Get-ExcelRawOrFallback $raw $summary 'ttxly' @('ValidationStatus')
        Set-ExcelDataCell $row 52 (Get-ExcelStatusLabel $status) 5
        Set-ExcelDataCell $row 53 (Get-ExcelValidationLabel $validation) 5

        $seller = [string](Get-ExcelRawOrFallback $raw $summary 'nbmst' @('SellerTaxCode'))
        $buyer = [string](Get-ExcelRawOrFallback $raw $summary 'nmmst' @('BuyerTaxCode'))
        $authority = [string](Get-ExcelRawOrFallback $raw $summary 'mhdon' @('TaxAuthorityCode'))
        $provider = [string](Get-ExcelRawOrFallback $raw $summary 'msttcgp' @('ProviderTaxCode'))
        $invoiceId = [string](Get-ExcelRawOrFallback $raw $summary 'id' @('InvoiceId'))
        $link = Get-ExcelLookupLink -ProviderTaxCode $provider -SellerTaxCode $seller -TaxAuthorityCode $authority -InvoiceId $invoiceId -CounterPartyTaxCodes @($seller, $buyer) -Mode summary
        $lookupCode = Get-ExcelLookupCode @(
            (Get-ExcelObjectValue $raw 'cttkhac' @()),
            (Get-ExcelObjectValue $raw 'ttkhac' @()),
            (Get-ExcelObjectValue $summary 'InvoiceLookupFields' @()),
            (Get-ExcelObjectValue $summary 'InvoiceAdditionalFields' @())
        )
        if ([string]::IsNullOrWhiteSpace($lookupCode)) { $lookupCode = Get-ExcelText $summary 'LookupCode' '' }
        if ([string]::IsNullOrWhiteSpace($lookupCode)) {
            $lookupCode = Get-ExcelProviderLookupCode -ProviderTaxCode $provider -SellerTaxCode $seller -TaxAuthorityCode $authority -InvoiceId $invoiceId
        }
        Set-ExcelDataCell $row 54 $provider 16
        $linkStyle = if ($link -match '^https?://') { 10 } else { 5 }
        Set-ExcelDataCell $row 55 $link $linkStyle $(if ($link -match '^https?://') { $link } else { '' })
        Set-ExcelDataCell $row 56 $lookupCode 16

        $statusNumber = 0
        [void][int]::TryParse([string]$status, [ref]$statusNumber)
        if ($statusNumber -ge 2 -and $statusNumber -le 6) {
            $chain = Get-ExcelText $summary 'RelatedChain' (Get-ExcelText $raw 'RelatedChain' '')
            $relatedInfo = Get-ExcelText $summary 'RelatedInfo' (Get-ExcelText $raw 'RelatedInfo' '')
            $originalType = Get-ExcelRawOrFallback $raw $summary 'lhdgoc' @('OriginalInvoiceType')
            $originalTemplate = Get-ExcelRawOrFallback $raw $summary 'khmshdgoc' @('OriginalTemplateCode')
            $originalSeries = Get-ExcelRawOrFallback $raw $summary 'khhdgoc' @('OriginalSeries')
            $originalNumber = Get-ExcelRawOrFallback $raw $summary 'shdgoc' @('OriginalNumber')
            $originalDate = Get-ExcelRawOrFallback $raw $summary 'tdlhdgoc' @('OriginalDate')
            $originalNote = Get-ExcelRawOrFallback $raw $summary 'gchdgoc' @('OriginalNote')
            Set-ExcelDataCell $row 58 ([string]$originalType) 16
            Set-ExcelDataCell $row 59 ([string]$originalTemplate) 16
            Set-ExcelDataCell $row 60 ([string]$originalSeries) 16
            Set-ExcelDataCell $row 61 ([string]$originalNumber) 16
            Set-ExcelDataCell $row 62 (ConvertTo-ExcelDate $originalDate) 6
            Set-ExcelDataCell $row 63 ([string]$originalNote) 5
            Set-ExcelDataCell $row 57 $chain 9
            Set-ExcelDataCell $row 64 $relatedInfo 9
            $row.Height = Get-ExcelRowHeight @($chain, $relatedInfo)
        }
        $result.Add([pscustomobject]@{ Direction = $direction; Row = $row })
    }
    return $result.ToArray()
}

# Read a scalar numeric field from a JSON-like object without converting an
# absent field to the string "0".
function Get-ExcelObjectNumber {
    param($Object, [string]$Name)
    $value = Get-ExcelObjectValue $Object $Name $null
    if ($null -eq $value) { return $null }
    return ConvertTo-ExcelNumber $value
}

function Get-ExcelCommonValues {
    param($Summary)
    return [ordered]@{
        InvoiceSeries = Get-ExcelText $Summary 'InvoiceSeries' ''
        InvoiceNumber = Get-ExcelText $Summary 'InvoiceNumber' ''
        InvoiceDate = ConvertTo-ExcelDate (Get-ExcelObjectValue $Summary 'InvoiceDate' $null)
        Currency = Get-ExcelText $Summary 'Currency' ''
        ExchangeRate = ConvertTo-ExcelNumber (Get-ExcelObjectValue $Summary 'ExchangeRate' $null)
        SellerName = Get-ExcelText $Summary 'SellerName' ''
        SellerTaxCode = Get-ExcelText $Summary 'SellerTaxCode' ''
        SellerAddress = Get-ExcelText $Summary 'SellerAddress' ''
        SellerSigningTime = ConvertTo-ExcelDate (Get-ExcelObjectValue $Summary 'SellerSigningTime' $null)
        TaxAuthorityCode = Get-ExcelText $Summary 'TaxAuthorityCode' ''
        TaxAuthoritySigningTime = ConvertTo-ExcelDate (Get-ExcelObjectValue $Summary 'TaxAuthoritySigningTime' $null)
        BuyerName = Get-ExcelText $Summary 'BuyerName' ''
        BuyerTaxCode = Get-ExcelText $Summary 'BuyerTaxCode' ''
        BuyerAddress = Get-ExcelText $Summary 'BuyerAddress' ''
        ProviderTaxCode = Get-ExcelText $Summary 'ProviderTaxCode' ''
        InvoiceId = Get-ExcelText $Summary 'InvoiceId' ''
        TotalTax = ConvertTo-ExcelNumber (Get-ExcelObjectValue $Summary 'TaxAmount' $null)
    }
}

function Add-ExcelCommonCells {
    param($Row, $Values, [switch]$ForXml)
    $columns = 1..14
    for ($index = 0; $index -lt $columns.Count; $index++) {
        $column = $columns[$index]
        $name = @('InvoiceSeries','InvoiceNumber','InvoiceDate','Currency','ExchangeRate','SellerName','SellerTaxCode','SellerAddress','SellerSigningTime','TaxAuthorityCode','TaxAuthoritySigningTime','BuyerName','BuyerTaxCode','BuyerAddress')[$index]
        $value = $Values[$name]
        if ($ForXml -and $column -eq 3) { $value = Get-ExcelText $Values $name '' }
        $style = if ($name -in @('InvoiceDate','SellerSigningTime','TaxAuthoritySigningTime')) { 6 } elseif ($name -eq 'ExchangeRate') { 7 } else { 5 }
        Set-ExcelDataCell $Row $column $value $style
    }
}

function Get-ExcelItemTaxRate {
    param($Detail)
    $type = Get-ExcelText $Detail 'TaxType' (Get-ExcelText $Detail 'ltsuat' '')
    if ($type -in @('KKKNT', 'KCT')) { return $type }
    return ConvertTo-ExcelTaxRate (Get-ExcelObjectValue $Detail 'TaxRate' $null)
}

function Get-ExcelDetailAmounts {
    param($Detail)
    $amount = ConvertTo-ExcelNumber (Get-ExcelObjectValue $Detail 'AmountBeforeTax' $null)
    $tax = ConvertTo-ExcelNumber (Get-ExcelObjectValue $Detail 'TaxAmount' $null)
    $withTax = ConvertTo-ExcelNumber (Get-ExcelObjectValue $Detail 'AmountWithTax' $null)
    $rate = Get-ExcelItemTaxRate $Detail
    $amountNumber = if ($amount -is [ValueType] -and $amount -isnot [bool]) { [double]$amount } else { $null }
    $taxNumber = if ($tax -is [ValueType] -and $tax -isnot [bool]) { [double]$tax } else { $null }
    $withTaxNumber = if ($withTax -is [ValueType] -and $withTax -isnot [bool]) { [double]$withTax } else { $null }
    $rateNumber = if ($rate -is [ValueType] -and $rate -isnot [bool]) { [double]$rate } else { $null }
    if (($null -eq $tax -or ($null -ne $taxNumber -and $taxNumber -eq 0)) -and $null -ne $amountNumber -and $null -ne $rateNumber) {
        $tax = $amountNumber * $rateNumber
        $taxNumber = [double]$tax
    }
    if (($null -eq $withTax -or ($null -ne $withTaxNumber -and $withTaxNumber -eq 0)) -and $null -ne $amountNumber -and $null -ne $taxNumber) {
        $withTax = $amountNumber + $taxNumber
    }
    return [pscustomobject]@{ Amount = $amount; Tax = $tax; WithTax = $withTax; Rate = $rate }
}

function New-ExcelDetailRows {
    param($Rows)
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($group in @(Get-ExcelDetailGroups $Rows)) {
        $summary = $group.Summary
        $values = Get-ExcelCommonValues $summary
        $provider = $values.ProviderTaxCode
        $link = Get-ExcelLookupLink -ProviderTaxCode $provider -SellerTaxCode $values.SellerTaxCode -TaxAuthorityCode $values.TaxAuthorityCode -InvoiceId $values.InvoiceId -CounterPartyTaxCodes @($values.SellerTaxCode, $values.BuyerTaxCode) -Mode detail
        $lookupCode = Get-ExcelText $summary 'LookupCode' ''
        if ([string]::IsNullOrWhiteSpace($lookupCode)) {
            $lookupCode = Get-ExcelLookupCode @(
                (Get-ExcelObjectValue $summary 'InvoiceLookupFields' @()),
                (Get-ExcelObjectValue $summary 'InvoiceAdditionalFields' @())
            )
        }
        if ([string]::IsNullOrWhiteSpace($lookupCode)) {
            $lookupCode = Get-ExcelProviderLookupCode -ProviderTaxCode $provider -SellerTaxCode $values.SellerTaxCode -TaxAuthorityCode $values.TaxAuthorityCode -InvoiceId $values.InvoiceId
        }
        $sumTax = 0.0
        $amounts = @()
        foreach ($detail in @($group.Details)) {
            $amountInfo = Get-ExcelDetailAmounts $detail
            $amounts += $amountInfo
            if ($amountInfo.Tax -is [ValueType]) { $sumTax += [double]$amountInfo.Tax }
        }
        $totalTax = $values.TotalTax
        for ($index = 0; $index -lt $group.Details.Count; $index++) {
            $detail = $group.Details[$index]
            $row = New-ExcelDataRow
            Add-ExcelCommonCells $row $values
            $amountInfo = $amounts[$index]
            $taxType = Get-ExcelText $detail 'TaxType' (Get-ExcelText $detail 'ltsuat' '')
            $taxRate = $amountInfo.Rate
            $rateStyle = if ($taxType -in @('KKKNT', 'KCT')) { 5 } else { 8 }
            Set-ExcelDataCell $row 15 (Get-ExcelText $detail 'LineNumber' '') 5
            Set-ExcelDataCell $row 16 (Get-ExcelText $detail 'Nature' '') 5
            Set-ExcelDataCell $row 17 (Get-ExcelText $detail 'ProductCode' '') 16
            Set-ExcelDataCell $row 18 (Get-ExcelText $detail 'Description' '') 5
            Set-ExcelDataCell $row 19 (Get-ExcelText $detail 'Unit' '') 5
            Set-ExcelDataCell $row 20 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'Quantity' $null)) 7
            Set-ExcelDataCell $row 21 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'UnitPrice' $null)) 7
            Set-ExcelDataCell $row 22 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'DiscountRate' $null)) 8
            Set-ExcelDataCell $row 23 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'DiscountAmount' $null)) 7
            Set-ExcelDataCell $row 24 $taxType 5
            Set-ExcelDataCell $row 25 $taxRate $rateStyle
            Set-ExcelDataCell $row 26 $amountInfo.Amount 7
            Set-ExcelDataCell $row 27 $amountInfo.Tax 7
            Set-ExcelDataCell $row 28 $amountInfo.WithTax 7
            if ($index -eq 0) {
                Set-ExcelDataCell $row 29 $totalTax 7
                $validationText = ''
                $validationStyle = 11
                if ($totalTax -is [ValueType]) {
                    if ([Math]::Abs(([double]$sumTax) - ([double]$totalTax)) -gt 0.000001) {
                        $validationText = 'Kiem tra lai tien thue: [{0}] <> [{1}]' -f $sumTax, $totalTax
                        $validationStyle = 12
                    }
                    else {
                        $validationText = 'Tien thue ok'
                    }
                }
                else { $validationStyle = 12 }
                Set-ExcelDataCell $row 30 $validationText $validationStyle
                Set-ExcelDataCell $row 31 $provider 16
                $linkStyle = if ($link -match '^https?://') { 10 } else { 5 }
                Set-ExcelDataCell $row 32 $link $linkStyle $(if ($link -match '^https?://') { $link } else { '' })
                Set-ExcelDataCell $row 33 $lookupCode 16
            }
            $result.Add([pscustomobject]@{ Direction = (Get-ExcelText $summary 'Direction' 'purchase').ToLowerInvariant(); Row = $row })
        }
    }
    return $result.ToArray()
}

function New-ExcelXmlRows {
    param($Rows)
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($group in @(Get-ExcelDetailGroups $Rows)) {
        $summary = $group.Summary
        $values = Get-ExcelCommonValues $summary
        $provider = $values.ProviderTaxCode
        $link = Get-ExcelLookupLink -ProviderTaxCode $provider -SellerTaxCode $values.SellerTaxCode -TaxAuthorityCode $values.TaxAuthorityCode -InvoiceId $values.InvoiceId -CounterPartyTaxCodes @($values.SellerTaxCode, $values.BuyerTaxCode) -Mode xml
        $lookupCode = Get-ExcelText $summary 'LookupCode' ''
        if ([string]::IsNullOrWhiteSpace($lookupCode)) {
            $lookupCode = Get-ExcelLookupCode @(
                (Get-ExcelObjectValue $summary 'InvoiceLookupFields' @()),
                (Get-ExcelObjectValue $summary 'InvoiceAdditionalFields' @())
            )
        }
        if ([string]::IsNullOrWhiteSpace($lookupCode)) {
            $lookupCode = Get-ExcelProviderLookupCode -ProviderTaxCode $provider -SellerTaxCode $values.SellerTaxCode -TaxAuthorityCode $values.TaxAuthorityCode -InvoiceId $values.InvoiceId
        }
        if ([string]::IsNullOrWhiteSpace($lookupCode)) { $lookupCode = 'Khong co ma tra cuu' }
        $direction = (Get-ExcelText $summary 'Direction' 'purchase').ToLowerInvariant()
        for ($index = 0; $index -lt $group.Details.Count; $index++) {
            $detail = $group.Details[$index]
            $row = New-ExcelDataRow
            $xmlDateText = Get-ExcelText $summary 'InvoiceDateText' (Get-ExcelText $values 'InvoiceDate' '')
            # XML luôn ghi NBan ở nhóm người bán/người mua theo tên node;
            # với sheet bán ra, template đã đảo nhãn để phản ánh hướng tra cứu.
            Set-ExcelDataCell $row 1 $values.InvoiceId 5
            Set-ExcelDataCell $row 2 $values.InvoiceSeries 5
            Set-ExcelDataCell $row 3 $values.InvoiceNumber 16
            Set-ExcelDataCell $row 4 $xmlDateText 5
            Set-ExcelDataCell $row 5 (Get-ExcelText $values 'Currency' '') 5
            Set-ExcelDataCell $row 6 (ConvertTo-ExcelNumber $values.ExchangeRate) 7
            Set-ExcelDataCell $row 7 $values.SellerName 5
            Set-ExcelDataCell $row 8 $values.SellerTaxCode 16
            Set-ExcelDataCell $row 9 $values.SellerAddress 5
            Set-ExcelDataCell $row 10 $values.SellerSigningTime 6
            Set-ExcelDataCell $row 11 $values.TaxAuthorityCode 16
            Set-ExcelDataCell $row 12 $values.TaxAuthoritySigningTime 6
            Set-ExcelDataCell $row 13 $values.BuyerName 5
            Set-ExcelDataCell $row 14 $values.BuyerTaxCode 16
            Set-ExcelDataCell $row 15 $values.BuyerAddress 5
            Set-ExcelDataCell $row 16 (Get-ExcelText $detail 'LineNumber' '') 5
            Set-ExcelDataCell $row 17 (Get-ExcelText $detail 'ProductCode' '') 16
            Set-ExcelDataCell $row 18 (Get-ExcelText $detail 'Description' '') 5
            Set-ExcelDataCell $row 19 (Get-ExcelText $detail 'Unit' '') 5
            Set-ExcelDataCell $row 20 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'Quantity' $null)) 7
            Set-ExcelDataCell $row 21 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'UnitPrice' $null)) 7
            Set-ExcelDataCell $row 22 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'DiscountRate' $null)) 8
            Set-ExcelDataCell $row 23 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'DiscountAmount' $null)) 7
            Set-ExcelDataCell $row 24 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'AmountBeforeTax' $null)) 7
            # parseXML writes the raw TSuat text in the XML sheet.
            Set-ExcelDataCell $row 25 (Get-ExcelText $detail 'TaxRate' '') 5
            Set-ExcelDataCell $row 26 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'TaxAmount' $null)) 7
            Set-ExcelDataCell $row 27 (ConvertTo-ExcelNumber (Get-ExcelObjectValue $detail 'AmountWithTax' $null)) 7
            if ($index -eq 0) {
                Set-ExcelDataCell $row 28 $values.TotalTax 7
                Set-ExcelDataCell $row 29 $provider 16
                $linkStyle = if ($link -match '^https?://') { 10 } else { 5 }
                Set-ExcelDataCell $row 30 $link $linkStyle $(if ($link -match '^https?://') { $link } else { '' })
                Set-ExcelDataCell $row 31 $lookupCode 16
                # Cột 32 ("Link tải hóa đơn gốc") được giữ theo template nguồn;
                # không tự bịa URL khi XML chỉ cung cấp dữ liệu cục bộ.
            }
            $result.Add([pscustomobject]@{ Direction = $direction; Row = $row })
        }
    }
    return $result.ToArray()
}

function Get-ExcelErrorKey {
    param($ErrorRow)
    return '{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}' -f (
        (Get-ExcelText $ErrorRow 'Direction' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'Source' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'Period' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'SellerTaxCode' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'InvoiceTemplate' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'InvoiceSeries' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'InvoiceNumber' '').ToUpperInvariant(),
        (Get-ExcelText $ErrorRow 'Stage' 'Tải XML').ToUpperInvariant()
    )
}

function Protect-ExcelErrorText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $safe = [string]$Text
    # Redact credentials embedded in a URL before redacting header names.
    $safe = [regex]::Replace($safe, '(?i)(https?://)[^/\s:@]+:[^@\s/]+@', '$1[REDACTED]@')
    $safe = [regex]::Replace($safe, '(?i)\bBearer\s+[A-Za-z0-9._~+/=-]+', 'Bearer [REDACTED]')
    $safe = [regex]::Replace($safe, '(?i)(Authorization\s*:\s*)[^\r\n]*', '$1[REDACTED]')
    $safe = [regex]::Replace($safe, '(?i)(proxy[_ -]?password|password|passwd|pwd)(\s*[:=]\s*)[^\s,;]+', '$1$2[REDACTED]')
    return $safe
}

function New-ExcelErrorRows {
    param($Rows)
    $unique = [ordered]@{}
    foreach ($error in @($Rows)) {
        if ($null -eq $error) { continue }
        $unique[(Get-ExcelErrorKey $error)] = $error
    }
    $result = New-Object System.Collections.Generic.List[object]
    $index = 0
    foreach ($error in $unique.Values) {
        $index++
        $row = New-ExcelDataRow
        $recordedAt = Get-ExcelObjectValue $error 'RecordedAt' $null
        if ($null -eq $recordedAt) { $recordedAt = [datetime]::Now }
        $invoiceDate = ConvertTo-ExcelDate (Get-ExcelObjectValue $error 'InvoiceDate' $null)
        $statusCode = Get-ExcelObjectValue $error 'StatusCode' $null
        $attempts = Get-ExcelObjectValue $error 'Attempts' $null
        $retryAfter = Get-ExcelObjectValue $error 'RetryAfterSeconds' $null
        $values = @(
            $index,
            $recordedAt,
            (Get-ExcelText $error 'Direction' ''),
            (Get-ExcelText $error 'Source' ''),
            (Get-ExcelText $error 'SellerTaxCode' ''),
            (Get-ExcelText $error 'InvoiceTemplate' ''),
            (Get-ExcelText $error 'InvoiceSeries' ''),
            (Get-ExcelText $error 'InvoiceNumber' ''),
            $invoiceDate,
            (Get-ExcelText $error 'Stage' 'Tải XML'),
            (Protect-ExcelErrorText (Get-ExcelText $error 'Endpoint' '')),
            $(if ($null -ne $statusCode -and [int]$statusCode -gt 0) { [int]$statusCode } else { '' }),
            (Protect-ExcelErrorText (Get-ExcelText $error 'Error' (Get-ExcelText $error 'Message' ''))),
            $(if ($null -ne $attempts) { [int]$attempts } else { '' }),
            $(if ($null -ne $retryAfter -and [int]$retryAfter -gt 0) { [int]$retryAfter } else { '' }),
            (Get-ExcelText $error 'FinalResult' 'Không tải được'),
            (Protect-ExcelErrorText (Get-ExcelText $error 'Note' ''))
        )
        for ($column = 1; $column -le $values.Count; $column++) {
            $style = 0
            if ($column -eq 2) { $style = 14 }
            elseif ($column -eq 9) { $style = 15 }
            elseif ($column -in @(5, 6, 7, 8, 12, 14, 15)) { $style = 5 }
            Set-ExcelDataCell $row $column $values[$column - 1] $style
        }
        $result.Add($row)
    }
    return $result.ToArray()
}

# Nạp bảng LinkTraCuu do người dùng cung cấp (LOOKUP_TABLE_XLSX).  Các dòng
# thêm vào được ghi lại vào sheet LinkTraCuu của workbook xuất và dùng để tạo
# link tra cứu, nên người dùng sửa bảng trong Excel là link trong workbook mới
# đổi theo.
function Import-ExcelLookupTable {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return 0 }
    $imported = Import-TaiHoaDonDienTuLookupTable -Path $Path
    $script:SourceStatusLabels = @(Get-TaiHoaDonDienTuStatusLabels)
    $script:SourceValidationLabels = @(Get-TaiHoaDonDienTuValidationLabels)
    return $imported
}

# Sheet LinkTraCuu: ghi lại nguyên bảng định tuyến tra cứu để người dùng thấy
# và tự sửa link.  Bảng này cũng chính là nguồn sinh link ở cột 55/56.
function New-ExcelLookupRows {
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($source in @(Get-TaiHoaDonDienTuLookupSheetRows)) {
        $row = New-ExcelDataRow
        for ($column = 0; $column -lt $script:TaiHoaDonDienTuLookupColumns.Count; $column++) {
            $letter = $script:TaiHoaDonDienTuLookupColumns[$column]
            $value = Get-TaiHoaDonDienTuLookupCell $source $letter
            if ([string]::IsNullOrWhiteSpace($value)) { continue }
            $index = $column + 1
            $style = if ($index -eq 2 -or $index -eq 3) { 16 } else { 5 }
            $hyperlink = ''
            if ($index -eq 4 -and $value -match '^https?://') {
                $style = 10
                $hyperlink = $value
            }
            Set-ExcelDataCell $row $index $value $style $hyperlink
        }
        $result.Add($row)
    }
    return $result.ToArray()
}

function Write-ExcelCell {
    param(
        [Parameter(Mandatory = $true)]$Writer,
        [Parameter(Mandatory = $true)][string]$Reference,
        $Value,
        [int]$Style = 0
    )
    Write-XmlStartElement $Writer 'c'
    $Writer.WriteAttributeString('r', $Reference)
    if ($Style -gt 0) { $Writer.WriteAttributeString('s', [string]$Style) }
    if ($null -eq $Value) {
        $Writer.WriteEndElement()
        return
    }
    if ($Value -is [datetime]) {
        $Writer.WriteAttributeString('t', 'n')
        Write-XmlStartElement $Writer 'v'
        $Writer.WriteString(([datetime]$Value).ToOADate().ToString('0.###############', [Globalization.CultureInfo]::InvariantCulture))
        $Writer.WriteEndElement()
    }
    elseif ($Value -is [bool]) {
        $Writer.WriteAttributeString('t', 'b')
        Write-XmlStartElement $Writer 'v'
        $Writer.WriteString($(if ($Value) { '1' } else { '0' }))
        $Writer.WriteEndElement()
    }
    elseif ($Value -is [byte] -or $Value -is [int16] -or $Value -is [int32] -or $Value -is [int64] -or $Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) {
        $Writer.WriteAttributeString('t', 'n')
        Write-XmlStartElement $Writer 'v'
        $Writer.WriteString([Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture))
        $Writer.WriteEndElement()
    }
    else {
        $text = [string]$Value
        $Writer.WriteAttributeString('t', 'inlineStr')
        Write-XmlStartElement $Writer 'is'
        Write-XmlStartElement $Writer 't'
        if ($text.Length -gt 0 -and ($text[0] -eq ' ' -or $text[$text.Length - 1] -eq ' ')) {
            $Writer.WriteAttributeString('xml', 'space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        }
        $Writer.WriteString($text)
        $Writer.WriteEndElement()
        $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Write-ExcelWorksheetXml {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Title,
        [string[]]$Headers,
        [object[]]$Rows,
        [double[]]$Widths,
        [int]$HeaderStyle = 2,
        [int]$TitleStyle = 1,
        [int]$HeaderRow = 2,
        [int]$DataStartRow = 3,
        [switch]$FreezeHeader,
        [switch]$AutoFilter
    )
    $rowItems = New-Object System.Collections.Generic.List[object]
    foreach ($item in @($Rows)) { if ($null -ne $item) { $rowItems.Add($item) } }
    $headerItems = @($Headers | Where-Object { $null -ne $_ })
    $lastColumn = Get-ExcelColumnName $headerItems.Count
    $lastRow = [Math]::Max($HeaderRow, $DataStartRow + $rowItems.Count - 1)
    $dimension = 'A1:{0}{1}' -f $lastColumn, $lastRow
    $writer = New-Utf8XmlWriter $Path
    $hyperlinks = New-Object System.Collections.Generic.List[object]
    try {
        Write-XmlStartElement $writer 'worksheet'
        $writer.WriteAttributeString('xmlns', 'r', $null, $script:RelationshipNamespace)

        Write-XmlStartElement $writer 'sheetPr'
        Write-XmlStartElement $writer 'pageSetUpPr'
        $writer.WriteAttributeString('fitToPage', '1')
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'dimension'
        $writer.WriteAttributeString('ref', $dimension)
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'sheetViews'
        Write-XmlStartElement $writer 'sheetView'
        $writer.WriteAttributeString('workbookViewId', '0')
        if ($FreezeHeader) {
            $topLeft = 'A{0}' -f ($HeaderRow + 1)
            Write-XmlStartElement $writer 'pane'
            $writer.WriteAttributeString('ySplit', [string]$HeaderRow)
            $writer.WriteAttributeString('topLeftCell', $topLeft)
            $writer.WriteAttributeString('activePane', 'bottomLeft')
            $writer.WriteAttributeString('state', 'frozen')
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'sheetFormatPr'
        $writer.WriteAttributeString('defaultRowHeight', '15')
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'cols'
        for ($column = 0; $column -lt $Widths.Count; $column++) {
            Write-XmlStartElement $writer 'col'
            $writer.WriteAttributeString('min', [string]($column + 1))
            $writer.WriteAttributeString('max', [string]($column + 1))
            $writer.WriteAttributeString('width', [Convert]::ToString($Widths[$column], [Globalization.CultureInfo]::InvariantCulture))
            $writer.WriteAttributeString('customWidth', '1')
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        Write-XmlStartElement $writer 'sheetData'
        if (-not [string]::IsNullOrWhiteSpace($Title)) {
            Write-XmlStartElement $writer 'row'
            $writer.WriteAttributeString('r', '1')
            $writer.WriteAttributeString('ht', '41.25')
            $writer.WriteAttributeString('customHeight', '1')
            Write-ExcelCell -Writer $writer -Reference 'A1' -Value $Title -Style $TitleStyle
            $writer.WriteEndElement()
        }
        Write-XmlStartElement $writer 'row'
        $writer.WriteAttributeString('r', [string]$HeaderRow)
        $writer.WriteAttributeString('ht', $(if ($HeaderStyle -eq 13) { '30' } else { '48' }))
        $writer.WriteAttributeString('customHeight', '1')
        for ($column = 0; $column -lt $headerItems.Count; $column++) {
            Write-ExcelCell -Writer $writer -Reference ((Get-ExcelColumnName ($column + 1)) + $HeaderRow) -Value $headerItems[$column] -Style $HeaderStyle
        }
        $writer.WriteEndElement()

        for ($rowIndex = 0; $rowIndex -lt $rowItems.Count; $rowIndex++) {
            $excelRow = $DataStartRow + $rowIndex
            $record = $rowItems[$rowIndex]
            Write-XmlStartElement $writer 'row'
            $writer.WriteAttributeString('r', [string]$excelRow)
            if ($null -ne $record.Height) {
                $writer.WriteAttributeString('ht', ([double]$record.Height).ToString('0.##', [Globalization.CultureInfo]::InvariantCulture))
                $writer.WriteAttributeString('customHeight', '1')
            }
            foreach ($column in ($record.Cells.Keys | Sort-Object)) {
                $cell = $record.Cells[$column]
                $reference = (Get-ExcelColumnName ([int]$column)) + $excelRow
                Write-ExcelCell -Writer $writer -Reference $reference -Value $cell.Value -Style ([int]$cell.Style)
                if (-not [string]::IsNullOrWhiteSpace([string]$cell.Hyperlink)) {
                    $hyperlinks.Add([pscustomobject]@{ Reference = $reference; Target = [string]$cell.Hyperlink })
                }
            }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()

        if ($AutoFilter) {
            $filterLastRow = [Math]::Max($HeaderRow, $lastRow)
            Write-XmlStartElement $writer 'autoFilter'
            $writer.WriteAttributeString('ref', ('A{0}:{1}{2}' -f $HeaderRow, $lastColumn, $filterLastRow))
            $writer.WriteEndElement()
        }
        if ($hyperlinks.Count -gt 0) {
            Write-XmlStartElement $writer 'hyperlinks'
            $relationshipId = 1
            foreach ($hyperlink in $hyperlinks) {
                Write-XmlStartElement $writer 'hyperlink'
                $hyperlinkWriter = $writer
                $hyperlinkWriter.WriteAttributeString('ref', $hyperlink.Reference)
                $hyperlinkWriter.WriteAttributeString('r', 'id', $script:RelationshipNamespace, ('rId{0}' -f $relationshipId))
                $hyperlinkWriter.WriteEndElement()
                $relationshipId++
            }
            $writer.WriteEndElement()
        }
        Write-XmlStartElement $writer 'printOptions'
        $writer.WriteAttributeString('horizontalCentered', '0')
        $writer.WriteEndElement()
        Write-XmlStartElement $writer 'pageMargins'
        $writer.WriteAttributeString('left', '0.25')
        $writer.WriteAttributeString('right', '0.25')
        $writer.WriteAttributeString('top', '0.5')
        $writer.WriteAttributeString('bottom', '0.5')
        $writer.WriteAttributeString('header', '0.2')
        $writer.WriteAttributeString('footer', '0.2')
        $writer.WriteEndElement()
        Write-XmlStartElement $writer 'pageSetup'
        $writer.WriteAttributeString('paperSize', '9')
        $writer.WriteAttributeString('orientation', 'landscape')
        $writer.WriteAttributeString('fitToWidth', '1')
        $writer.WriteAttributeString('fitToHeight', '0')
        $writer.WriteEndElement()
        $writer.WriteEndElement()
    }
    finally { $writer.Dispose() }
    return $hyperlinks.ToArray()
}

function Write-ExcelSheetRelationships {
    param([string]$Path, $Hyperlinks)
    $items = @($Hyperlinks)
    if ($items.Count -eq 0) { return }
    $writer = New-Utf8XmlWriter $Path
    try {
        [void]$writer.WriteStartElement('Relationships', $script:PackageRelationshipNamespace)
        $index = 1
        foreach ($hyperlink in $items) {
            [void]$writer.WriteStartElement('Relationship', $script:PackageRelationshipNamespace)
            $writer.WriteAttributeString('Id', ('rId{0}' -f $index))
            $writer.WriteAttributeString('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')
            $writer.WriteAttributeString('Target', [string]$hyperlink.Target)
            $writer.WriteAttributeString('TargetMode', 'External')
            [void]$writer.WriteEndElement()
            $index++
        }
        [void]$writer.WriteEndElement()
    }
    finally { $writer.Dispose() }
}

function Write-ExcelStylesXml {
    param([string]$Path)
    $styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="165" formatCode="dd/mm/yyyy"/>
    <numFmt numFmtId="166" formatCode="dd/mm/yyyy hh:mm:ss"/>
    <numFmt numFmtId="167" formatCode="_(* #,##0_);_(* \(#,##0\);_(* &quot;-&quot;??_);_(@_)"/>
  </numFmts>
  <fonts count="7">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
    <font><sz val="11"/><name val="Consolas"/><family val="3"/></font>
    <font><b/><sz val="18"/><color rgb="FF7030A0"/><name val="Consolas"/><family val="3"/></font>
    <font><u/><color rgb="FF0563C1"/><sz val="11"/><name val="Consolas"/><family val="3"/></font>
    <font><color rgb="FF0000FF"/><sz val="11"/><name val="Consolas"/><family val="3"/></font>
    <font><color rgb="FFFF0000"/><sz val="11"/><name val="Consolas"/><family val="3"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE0E1AD"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="17">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="49" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="167" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="9" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
    <xf numFmtId="166" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
    <xf numFmtId="49" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>
</styleSheet>
'@
    Write-TextFileUtf8NoBom $Path $styles
}

# Đường dẫn part trong gói OOXML phải ghép từng phần tử riêng.  Chuỗi kiểu
# 'xl\worksheets\sheet1.xml' chỉ đúng trên Windows: Join-Path trên macOS/Linux coi
# '\' là dấu phân cách, nên '_rels\.rels' trở thành file ẩn '.rels'.  Get-ChildItem
# mặc định bỏ qua file ẩn, gói xuất ra thiếu hẳn '_rels/.rels' và Excel báo
# "Sorry, we couldn't open your workbook".
function Join-ExcelPackagePath {
    param([Parameter(Mandatory = $true)][string]$Root, [Parameter(Mandatory = $true)][string[]]$Segments)
    $path = $Root
    foreach ($segment in $Segments) { $path = Join-Path $path $segment }
    return $path
}

# Danh sách part bắt buộc của một gói SpreadsheetML hợp lệ.
$script:RequiredPackageParts = @(
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/workbook.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml'
)

function Add-ExcelPackagePart {
    param($Parts, [string]$Name, [string]$Path)
    [void]$Parts.Add([pscustomobject]@{ Name = $Name; Path = $Path })
}

# Kiểm tra gói vừa ghi trước khi thay thế file cũ, để một gói hỏng bao giờ
# cũng không thành file người dùng phải mở.
function Assert-ExcelPackage {
    param([Parameter(Mandatory = $true)][string]$Path, [object[]]$Parts)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entries = @($archive.Entries | ForEach-Object { [string]$_.FullName })
    }
    finally { $archive.Dispose() }
    $expected = @($script:RequiredPackageParts) + @($Parts | ForEach-Object { [string]$_.Name })
    $missing = @($expected | Where-Object { $entries -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw ('Gói Excel không hợp lệ, thiếu thành phần: ' + (($missing | Select-Object -Unique) -join ', '))
    }
}

function Export-InvoiceWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$SummaryRows,
        [Parameter(Mandatory = $true)]$DetailRows,
        [Parameter(Mandatory = $true)]$ErrorRows,
        [switch]$Overwrite,
        # Workbook .xlsx có sheet 'LinkTraCuu' do người dùng sửa; các dòng
        # MST/link trong đó được nối vào bảng tra cứu dùng cho file này.
        [string]$LookupTablePath = ''
    )

    if ((Test-Path -LiteralPath $Path) -and -not $Overwrite) {
        throw "File Excel đã tồn tại: $Path. Đặt OVERWRITE_OUTPUT=true để ghi đè."
    }
    $parent = Split-Path -Parent $Path
    if ([string]::IsNullOrWhiteSpace($parent)) { $parent = (Get-Location).Path }
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    $importedLookupRows = 0
    if (-not [string]::IsNullOrWhiteSpace($LookupTablePath)) {
        $importedLookupRows = Import-ExcelLookupTable -Path $LookupTablePath
    }

    $summaryData = New-ExcelSummaryRows $SummaryRows
    $detailData = New-ExcelDetailRows $DetailRows
    $xmlData = New-ExcelXmlRows $DetailRows
    $errorData = New-ExcelErrorRows $ErrorRows
    $purchaseSummary = @($summaryData | Where-Object Direction -eq 'purchase' | ForEach-Object Row)
    $soldSummary = @($summaryData | Where-Object Direction -eq 'sold' | ForEach-Object Row)
    $purchaseDetail = @($detailData | Where-Object Direction -eq 'purchase' | ForEach-Object Row)
    $soldDetail = @($detailData | Where-Object Direction -eq 'sold' | ForEach-Object Row)
    $purchaseXml = @($xmlData | Where-Object Direction -eq 'purchase' | ForEach-Object Row)
    $soldXml = @($xmlData | Where-Object Direction -eq 'sold' | ForEach-Object Row)

    $sheetDefinitions = @(
        [pscustomobject]@{ Name='TongHopHD_Mua'; Title='HÓA ĐƠN MUA'; Headers=$script:SourceSummaryHeaders; Rows=$purchaseSummary; Widths=$script:SourceSummaryWidths; HeaderStyle=2; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='ChiTietHD_Mua'; Title='HÓA ĐƠN MUA - CHI TIẾT'; Headers=$script:SourceDetailHeaders; Rows=$purchaseDetail; Widths=$script:SourceDetailWidths; HeaderStyle=3; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='ChiTietHD_Mua_XML'; Title='HÓA ĐƠN MUA - CHI TIẾT - XML'; Headers=$script:SourceXmlHeadersPurchase; Rows=$purchaseXml; Widths=$script:SourceXmlWidths; HeaderStyle=4; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='TongHopHD_Ban'; Title='HÓA ĐƠN BÁN'; Headers=$script:SourceSummaryHeaders; Rows=$soldSummary; Widths=$script:SourceSummaryWidths; HeaderStyle=2; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='ChiTietHD_Ban'; Title='HÓA ĐƠN BÁN - CHI TIẾT'; Headers=$script:SourceDetailHeaders; Rows=$soldDetail; Widths=$script:SourceDetailWidths; HeaderStyle=3; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='ChiTietHD_Ban_XML'; Title='HÓA ĐƠN BÁN - CHI TIẾT - XML'; Headers=$script:SourceXmlHeadersSold; Rows=$soldXml; Widths=$script:SourceXmlWidths; HeaderStyle=4; HeaderRow=2; DataStartRow=3; Freeze=$true; Filter=$false },
        [pscustomobject]@{ Name='BaoCao_LoiTaiHD'; Title=''; Headers=$script:SourceErrorHeaders; Rows=$errorData; Widths=$script:SourceErrorWidths; HeaderStyle=13; HeaderRow=1; DataStartRow=2; Freeze=$true; Filter=$true },
        [pscustomobject]@{ Name='LinkTraCuu'; Title=''; Headers=$script:SourceLookupHeaders; Rows=(New-ExcelLookupRows); Widths=$script:SourceLookupWidths; HeaderStyle=2; HeaderRow=1; DataStartRow=2; Freeze=$true; Filter=$false }
    )

    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('hddt-xlsx-' + [guid]::NewGuid().ToString('N'))
    $stagingPath = Join-Path $parent ('.' + [IO.Path]::GetFileName($Path) + '.hddt-' + [guid]::NewGuid().ToString('N') + '.tmp')
    $backupPath = Join-Path $parent ('.' + [IO.Path]::GetFileName($Path) + '.hddt-' + [guid]::NewGuid().ToString('N') + '.bak')
    New-Item -ItemType Directory -Path (Join-ExcelPackagePath $tempRoot @('_rels')) -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-ExcelPackagePath $tempRoot @('xl', '_rels')) -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-ExcelPackagePath $tempRoot @('xl', 'worksheets', '_rels')) -Force | Out-Null
    # Gói được dựng từ danh sách part tường minh thay vì quét thư mục tạm:
    # không còn phụ thuộc vào quy tắc file ẩn của hệ điều hành.
    $packageParts = New-Object System.Collections.Generic.List[object]
    try {
        for ($index = 0; $index -lt $sheetDefinitions.Count; $index++) {
            $sheetNumber = $index + 1
            $sheet = $sheetDefinitions[$index]
            $sheetPartName = 'xl/worksheets/sheet{0}.xml' -f $sheetNumber
            $sheetPath = Join-ExcelPackagePath $tempRoot @('xl', 'worksheets', ('sheet{0}.xml' -f $sheetNumber))
            $hyperlinks = Write-ExcelWorksheetXml -Path $sheetPath -Title $sheet.Title -Headers $sheet.Headers -Rows $sheet.Rows -Widths $sheet.Widths -HeaderStyle $sheet.HeaderStyle -HeaderRow $sheet.HeaderRow -DataStartRow $sheet.DataStartRow -FreezeHeader:$sheet.Freeze -AutoFilter:$sheet.Filter
            Add-ExcelPackagePart $packageParts $sheetPartName $sheetPath
            if (@($hyperlinks).Count -gt 0) {
                $sheetRelsPartName = 'xl/worksheets/_rels/sheet{0}.xml.rels' -f $sheetNumber
                $sheetRelsPath = Join-ExcelPackagePath $tempRoot @('xl', 'worksheets', '_rels', ('sheet{0}.xml.rels' -f $sheetNumber))
                Write-ExcelSheetRelationships -Path $sheetRelsPath -Hyperlinks $hyperlinks
                Add-ExcelPackagePart $packageParts $sheetRelsPartName $sheetRelsPath
            }
        }

        $sheetOverrides = for ($index = 1; $index -le $sheetDefinitions.Count; $index++) {
            '<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $index
        }
        $contentTypes = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="$script:ContentTypeNamespace"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>$($sheetOverrides -join '')</Types>
"@
        $contentTypesPath = Join-ExcelPackagePath $tempRoot @('[Content_Types].xml')
        Write-TextFileUtf8NoBom $contentTypesPath $contentTypes
        Add-ExcelPackagePart $packageParts '[Content_Types].xml' $contentTypesPath
        $rootRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$script:PackageRelationshipNamespace"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
"@
        $rootRelsPath = Join-ExcelPackagePath $tempRoot @('_rels', '.rels')
        Write-TextFileUtf8NoBom $rootRelsPath $rootRels
        Add-ExcelPackagePart $packageParts '_rels/.rels' $rootRelsPath
        $sheetNodes = for ($index = 0; $index -lt $sheetDefinitions.Count; $index++) {
            '<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f $sheetDefinitions[$index].Name, ($index + 1)
        }
        $workbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="$script:SpreadsheetNamespace" xmlns:r="$script:RelationshipNamespace"><bookViews><workbookView activeTab="0"/></bookViews><sheets>$($sheetNodes -join '')</sheets></workbook>
"@
        $workbookPath = Join-ExcelPackagePath $tempRoot @('xl', 'workbook.xml')
        Write-TextFileUtf8NoBom $workbookPath $workbookXml
        Add-ExcelPackagePart $packageParts 'xl/workbook.xml' $workbookPath
        $workbookRelationships = for ($index = 0; $index -lt $sheetDefinitions.Count; $index++) {
            '<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f ($index + 1)
        }
        $stylesRelationshipId = $sheetDefinitions.Count + 1
        $workbookRelsXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="$script:PackageRelationshipNamespace">$($workbookRelationships -join '')<Relationship Id="rId$stylesRelationshipId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>
"@
        $workbookRelsPath = Join-ExcelPackagePath $tempRoot @('xl', '_rels', 'workbook.xml.rels')
        Write-TextFileUtf8NoBom $workbookRelsPath $workbookRelsXml
        Add-ExcelPackagePart $packageParts 'xl/_rels/workbook.xml.rels' $workbookRelsPath
        $stylesPath = Join-ExcelPackagePath $tempRoot @('xl', 'styles.xml')
        Write-ExcelStylesXml $stylesPath
        Add-ExcelPackagePart $packageParts 'xl/styles.xml' $stylesPath

        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $outputStream = [IO.File]::Open($stagingPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $archive = New-Object IO.Compression.ZipArchive($outputStream, [IO.Compression.ZipArchiveMode]::Create, $false)
            try {
                foreach ($part in $packageParts) {
                    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $part.Path, $part.Name, [IO.Compression.CompressionLevel]::Optimal)
                }
            }
            finally { $archive.Dispose() }
        }
        finally { $outputStream.Dispose() }

        # Chỉ đổi tên file sau khi đã đóng hoàn toàn VÀ đã kiểm tra gói còn đủ
        # thành phần bắt buộc. Nếu file đang mở, thay thế thất bại hoặc gói
        # hỏng, bản cũ vẫn còn nguyên.
        Assert-ExcelPackage -Path $stagingPath -Parts $packageParts.ToArray()
        if (Test-Path -LiteralPath $Path) {
            [IO.File]::Replace($stagingPath, $Path, $backupPath)
            if (Test-Path -LiteralPath $backupPath) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue }
        }
        else {
            [IO.File]::Move($stagingPath, $Path)
        }
    }
    finally {
        $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
        $systemTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        if ($resolvedTempRoot.StartsWith($systemTempRoot, [StringComparison]::OrdinalIgnoreCase) -and (Test-Path -LiteralPath $resolvedTempRoot)) {
            Remove-Item -LiteralPath $resolvedTempRoot -Recurse -Force
        }
        if (Test-Path -LiteralPath $stagingPath) { Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $backupPath) { Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue }
    }

    # Báo lại để người dùng thấy ngay bao nhiêu hóa đơn tìm được link tra cứu;
    # phần còn lại là hóa đơn không có nhà cung cấp T-VAN trong bảng LinkTraCuu.
    $summaryRowList = @($purchaseSummary) + @($soldSummary)
    $linkedCount = 0
    foreach ($summaryRow in $summaryRowList) {
        if (-not $summaryRow.Cells.ContainsKey(55)) { continue }
        if (([string]$summaryRow.Cells[55].Value) -match '^https?://') { $linkedCount++ }
    }
    return [pscustomobject]@{
        Path = $Path
        SummaryRows = $summaryRowList.Count
        DetailRows = @($purchaseDetail).Count + @($soldDetail).Count
        LinksResolved = $linkedCount
        LinksMissing = $summaryRowList.Count - $linkedCount
        LookupSheetRows = @(Get-TaiHoaDonDienTuLookupSheetRows).Count
        ImportedLookupRows = $importedLookupRows
    }
}
