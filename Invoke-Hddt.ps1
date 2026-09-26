[CmdletBinding()]
param(
    [string]$EnvFile,
    [switch]$Interactive
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($EnvFile)) { $EnvFile = Join-Path $PSScriptRoot '.env' }

. (Join-Path $PSScriptRoot 'src\Logging.ps1')
. (Join-Path $PSScriptRoot 'src\Config.ps1')
. (Join-Path $PSScriptRoot 'src\Http.ps1')
. (Join-Path $PSScriptRoot 'src\InvoiceApi.ps1')
. (Join-Path $PSScriptRoot 'src\XmlParser.ps1')
. (Join-Path $PSScriptRoot 'src\ExcelExporter.ps1')
. (Join-Path $PSScriptRoot 'src\Login.ps1')

function Set-HddtObjectValue {
    param($Object, [string]$Name, $Value)
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        Add-Member -InputObject $Object -NotePropertyName $Name -NotePropertyValue $Value
    }
    else {
        $property.Value = $Value
    }
}

function Merge-HddtParsedSummary {
    param($Target, $Parsed)
    if ($null -eq $Target -or $null -eq $Parsed) { return }
    foreach ($property in $Parsed.PSObject.Properties) {
        if ($property.Name -in @('Direction', 'Source', 'XmlFile', 'GdtIndex', 'Status', 'RelatedChain', 'RelatedInfo', 'OriginalInvoiceType', 'OriginalTemplateCode', 'OriginalSeries', 'OriginalNumber', 'OriginalDate', 'OriginalNote')) { continue }
        $current = Get-ObjectValue $Target $property.Name $null
        if ($null -eq $current -or ($current -is [string] -and [string]::IsNullOrWhiteSpace($current))) {
            Set-HddtObjectValue $Target $property.Name $property.Value
        }
    }
    if ([string]::IsNullOrWhiteSpace([string](Get-ObjectValue $Target 'XmlFile' ''))) {
        Set-HddtObjectValue $Target 'XmlFile' $Parsed.XmlFile
    }
}

# Quy ước log INFO: [THẺ GIAI ĐOẠN] nội dung | chỉ số | thời gian.
# Các thẻ: CẤU HÌNH, ĐĂNG NHẬP, DANH SÁCH, TẢI XML, LIÊN QUAN, XUẤT FILE, KẾT QUẢ.
Initialize-HddtConsole
$loggingStarted = $false

if (Register-HddtStopHandler) {
    Write-Host 'Ctrl+C: dừng an toàn sau request hiện tại (nhấn lần thứ hai để dừng ngay).' -ForegroundColor DarkGray
}

try {
    $config = Get-HddtConfig -EnvFile $EnvFile -RepositoryRoot $PSScriptRoot -Interactive:$Interactive
    New-Item -ItemType Directory -Path $config.OutputDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $config.XmlDirectory -Force | Out-Null
    foreach ($direction in $config.Directions) {
        New-Item -ItemType Directory -Path (Join-Path $config.XmlDirectory $direction) -Force | Out-Null
    }
    $logFile = Start-HddtLogging -Directory $config.LogDirectory -Level $config.LogLevel -ToFile $config.LogToFile -RunName 'download'
    $loggingStarted = $true

    $familyNames = @()
    if ($config.IncludeRegular) { $familyNames += 'query' }
    if ($config.IncludeSco) { $familyNames += 'sco-query' }
    Write-HddtLog INFO 'Bắt đầu HDDT Downloader.'
    Write-HddtLog INFO ('[CẤU HÌNH] Phạm vi: {0} | {1:dd/MM/yyyy} - {2:dd/MM/yyyy} | nguồn: {3}' -f ($config.Directions -join '+'), $config.FromDate, $config.ToDate, ($familyNames -join '+'))
    $adaptiveText = if ($config.AdaptiveThrottle) { 'bật' } else { 'tắt' }
    $xmlModeText = if ($config.RedownloadXml) { 'tải lại từ đầu' } else { 'tiếp tục/tái sử dụng' }
    Write-HddtLog INFO ('[CẤU HÌNH] Mạng: tuần tự | page size {0} | giãn cách {1} ms | tự điều tiết {2} | retry {3} | timeout {4}s' -f $config.PageSize, $config.RequestDelayMs, $adaptiveText, $config.MaxRetries, $config.HttpTimeoutSeconds)
    if ($null -ne $config.ProxyUri) {
        Write-HddtLog INFO ('[CẤU HÌNH] Proxy: {0}:{1}' -f $config.ProxyUri.Host, $config.ProxyUri.Port)
    }
    else {
        Write-HddtLog INFO '[CẤU HÌNH] Proxy: không dùng.'
    }
    Write-HddtLog INFO ('[CẤU HÌNH] Chế độ XML: {0} | liên quan: {1}.' -f $xmlModeText, $(if ($config.FetchRelated) { 'gọi API' } else { 'chỉ dữ liệu danh sách' }))
    Write-HddtLog INFO ('[CẤU HÌNH] Đầu ra: {0}' -f $config.OutputWorkbook)
    if ($config.LogToFile) { Write-HddtLog INFO ('[CẤU HÌNH] Nhật ký: {0}' -f $logFile) }

    Write-HddtLog INFO ('[ĐĂNG NHẬP] Bắt đầu đăng nhập tài khoản {0}: tự lấy và nhận diện CAPTCHA.' -f $config.Username)
    $config.Token = Invoke-GdtLogin -Config $config
    Write-HddtLog INFO ('[ĐĂNG NHẬP] Thành công | token chỉ giữ trong bộ nhớ phiên | {0}' -f (Get-HddtElapsedText))

    $allInvoices = New-Object System.Collections.Generic.List[object]
    $indexErrors = New-Object System.Collections.Generic.List[object]
    foreach ($direction in $config.Directions) {
        if (Test-HddtStopRequested) {
            Write-HddtLog WARN '[DANH SÁCH] Đã yêu cầu dừng; không lấy thêm danh sách hóa đơn.'
            break
        }
        $beforeCount = $allInvoices.Count
        Write-HddtLog INFO ('[DANH SÁCH] Bắt đầu lấy {0}...' -f $direction)
        try {
            $items = @(Get-GdtInvoiceIndex -Config $config -Direction $direction)
        }
        catch {
            if (Test-HddtStopRequested) {
                Write-HddtLog WARN ('[DANH SÁCH] Dừng giữa chừng {0}: {1}' -f $direction, $_.Exception.Message)
                break
            }
            throw
        }
        foreach ($item in $items) { $allInvoices.Add($item) }
        foreach ($indexError in @(Get-GdtIndexErrors)) { $indexErrors.Add($indexError) }
        Write-HddtLog INFO ('[DANH SÁCH] {0}: +{1} hóa đơn | tổng {2} | chạy {3}.' -f $direction, ($allInvoices.Count - $beforeCount), $allInvoices.Count, (Get-HddtElapsedText))
    }

    Write-HddtLog INFO ('[TẢI XML] Tổng cộng {0} hóa đơn; bắt đầu tải và parse.' -f $allInvoices.Count)
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $summaryBindings = New-Object System.Collections.Generic.List[object]
    $detailRows = New-Object System.Collections.Generic.List[object]
    $errorRows = New-Object System.Collections.Generic.List[object]
    foreach ($indexError in $indexErrors) { $errorRows.Add($indexError) }
    $current = 0

    foreach ($invoice in $allInvoices) {
        if (Test-HddtStopRequested) {
            Write-HddtLog WARN ('[TẢI XML] Đã yêu cầu dừng; giữ lại {0} hóa đơn đã tải.' -f $summaryRows.Count)
            break
        }
        $current++
        $label = Get-InvoiceLabel -Invoice $invoice
        $detailBefore = $detailRows.Count
        # Ghi dòng tổng hợp ngay khi nhận được item danh sách. Nếu XML lỗi,
        # người dùng vẫn thấy hóa đơn và lỗi tương ứng trong workbook.
        $summary = [pscustomobject]@{
            Direction = $invoice.Direction
            Source = $invoice.Source
            XmlFile = ''
            InvoiceId = ''
            TemplateCode = $invoice.InvoiceTemplate
            InvoiceSeries = $invoice.InvoiceSeries
            InvoiceNumber = $invoice.InvoiceNumber
            InvoiceDate = $invoice.InvoiceDate
            SellerTaxCode = $invoice.SellerTaxCode
            Status = $invoice.Status
            ValidationStatus = $invoice.ValidationStatus
            GdtIndex = $invoice.GdtIndex
            RelatedChain = $invoice.RelatedChain
            RelatedInfo = ''
            OriginalInvoiceType = $invoice.OriginalInvoiceType
            OriginalTemplateCode = $invoice.OriginalTemplateCode
            OriginalSeries = $invoice.OriginalSeries
            OriginalNumber = $invoice.OriginalNumber
            OriginalDate = $invoice.OriginalDate
            OriginalNote = $invoice.OriginalNote
        }
        $summaryRows.Add($summary)
        $summaryBindings.Add([pscustomobject]@{ Summary = $summary; Invoice = $invoice })
        $stage = 'Tải XML'
        try {
            Write-HddtLog DEBUG ('Bắt đầu hóa đơn {0}/{1}: {2} [{3}]' -f $current, $allInvoices.Count, $label, $invoice.Source)
            $xmlFiles = @(Save-GdtInvoiceXml -Config $config -Invoice $invoice)
            $firstXml = $true
            foreach ($xmlFile in $xmlFiles) {
                $stage = 'Parse XML'
                $parsed = ConvertFrom-InvoiceXml -Path $xmlFile -Direction $invoice.Direction -Source $invoice.Source
                if ($firstXml) {
                    Merge-HddtParsedSummary -Target $summary -Parsed $parsed.Summary
                    $firstXml = $false
                }
                foreach ($row in $parsed.Details) { $detailRows.Add($row) }
            }

            if (($current % $config.ProgressEvery) -eq 0 -or $current -eq $allInvoices.Count) {
                $percent = if ($allInvoices.Count -eq 0) { 100 } else { [Math]::Floor(($current * 100.0) / $allInvoices.Count) }
                Write-HddtLog INFO ('[TẢI XML] [{0}/{1} | {2}%] {3} | XML {4} | +{5} dòng chi tiết' -f $current, $allInvoices.Count, $percent, $label, $xmlFiles.Count, ($detailRows.Count - $detailBefore))
            }
        }
        catch {
            if (Test-HddtStopRequested) {
                Write-HddtLog WARN ('[TẢI XML] Dừng khi tải {0}: {1}' -f $label, $_.Exception.Message)
                break
            }
            $errorRows.Add([pscustomobject]@{
                Direction = $invoice.Direction
                Source = $invoice.Source
                SellerTaxCode = $invoice.SellerTaxCode
                InvoiceTemplate = $invoice.InvoiceTemplate
                InvoiceSeries = $invoice.InvoiceSeries
                InvoiceNumber = $invoice.InvoiceNumber
                InvoiceDate = $invoice.InvoiceDate
                Stage = $stage
                Endpoint = Get-GdtLastRequestUri
                StatusCode = Get-GdtLastStatusCode
                Attempts = Get-GdtLastRequestAttempts
                RetryAfterSeconds = Get-GdtLastRetryAfterSeconds
                RecordedAt = [datetime]::Now
                Error = $_.Exception.Message
                FinalResult = 'Không tải được'
                Note = ''
            })
            Write-HddtLog WARN ('[TẢI XML] Không tải được {0} ({1}/{2}): {3}' -f $label, $current, $allInvoices.Count, $_.Exception.Message)
        }
    }

    # Chuỗi hóa đơn liên quan và thông tin liên quan (tương ứng
    # ProcessRelatedInvoiceApis trong VBA). Ghi thẳng vào dòng tổng hợp.
    $relatedColumns = @('RelatedChain', 'OriginalInvoiceType', 'OriginalTemplateCode', 'OriginalSeries', 'OriginalNumber', 'OriginalDate', 'OriginalNote', 'RelatedInfo')
    $relatedCandidates = @($summaryBindings | Where-Object { Test-GdtRelatedInvoice -Invoice $_.Invoice })
    $relatedHandled = 0
    if ($relatedCandidates.Count -gt 0) {
        if ($config.FetchRelated) {
            Write-HddtLog INFO ('[LIÊN QUAN] {0} hóa đơn trạng thái 2-6; lấy chuỗi liên quan và thông tin liên quan.' -f $relatedCandidates.Count)
        }
        else {
            Write-HddtLog INFO ('[LIÊN QUAN] {0} hóa đơn trạng thái 2-6; chỉ ghi thông tin hóa đơn gốc từ danh sách (FETCH_RELATED=false).' -f $relatedCandidates.Count)
        }
    }
    foreach ($binding in $relatedCandidates) {
        if (Test-HddtStopRequested) {
            Write-HddtLog WARN ('[LIÊN QUAN] Dừng giữa chừng; đã xử lý {0}/{1} hóa đơn.' -f $relatedHandled, $relatedCandidates.Count)
            break
        }
        try {
            if ($config.FetchRelated) {
                $relation = Get-GdtInvoiceRelation -Config $config -Invoice $binding.Invoice
            }
            else {
                $relation = Get-GdtInvoiceRelation -Config $config -Invoice $binding.Invoice -ListOnly
            }
        }
        catch {
            if (Test-HddtStopRequested) {
                Write-HddtLog WARN ('[LIÊN QUAN] Dừng khi lấy thông tin liên quan: {0}' -f $_.Exception.Message)
                break
            }
            throw
        }
        if ($null -eq $relation) { continue }
        foreach ($columnName in $relatedColumns) {
            Set-HddtObjectValue -Object $binding.Summary -Name $columnName `
                -Value ([string](Get-ObjectValue $relation $columnName ''))
        }
        $relatedHandled++
    }
    if ($relatedHandled -gt 0) {
        Write-HddtLog INFO ('[LIÊN QUAN] Hoàn tất: {0}/{1} hóa đơn | chạy {2}.' -f $relatedHandled, $relatedCandidates.Count, (Get-HddtElapsedText))
    }

    $hasData = ($summaryRows.Count -gt 0 -or $detailRows.Count -gt 0 -or $errorRows.Count -gt 0)
    $stopRequested = Test-HddtStopRequested
    $exported = $false

    if ($stopRequested -and -not $hasData) {
        Write-HddtLog WARN '[XUẤT FILE] Đã dừng khi chưa tải được dữ liệu nào; không ghi file Excel để tránh đè kết quả cũ.'
    }
    else {
        if ($stopRequested) {
            Write-HddtLog INFO '[XUẤT FILE] Xuất phần dữ liệu đã tải trước khi dừng...'
        }
        else {
            Write-HddtLog INFO ('[XUẤT FILE] Tạo workbook: tổng hợp {0} | chi tiết {1} | lỗi {2}.' -f $summaryRows.Count, $detailRows.Count, $errorRows.Count)
        }
        $export = Export-InvoiceWorkbook -Path $config.OutputWorkbook -SummaryRows ($summaryRows.ToArray()) -DetailRows ($detailRows.ToArray()) -ErrorRows ($errorRows.ToArray()) -Overwrite:$config.OverwriteOutput -LookupTablePath $config.LookupTableXlsx
        $exported = $true
        if ($export.LinksMissing -gt 0) {
            Write-HddtLog INFO ('[XUẤT FILE] Link tra cứu: {0}/{1} hóa đơn có link; {2} hóa đơn không có nhà cung cấp trong bảng LinkTraCuu.' -f $export.LinksResolved, $export.SummaryRows, $export.LinksMissing)
        }
        else {
            Write-HddtLog INFO ('[XUẤT FILE] Link tra cứu: {0}/{1} hóa đơn có link.' -f $export.LinksResolved, $export.SummaryRows)
        }
        if ($export.ImportedLookupRows -gt 0) {
            Write-HddtLog INFO ('[XUẤT FILE] Bảng tra cứu: nạp thêm {0} dòng từ {1}.' -f $export.ImportedLookupRows, $config.LookupTableXlsx)
        }
    }

    if ($stopRequested) {
        Write-HddtLog WARN ('[KẾT QUẢ] Đã dừng theo yêu cầu | chạy {0} | dữ liệu đã tải trước khi dừng vẫn được giữ lại.' -f (Get-HddtElapsedText))
    }
    else {
        Write-HddtLog INFO ('[KẾT QUẢ] Hoàn tất | chạy {0} | file: {1}' -f (Get-HddtElapsedText), $config.OutputWorkbook)
    }
    if ($exported) {
        Write-HddtLog INFO ('[KẾT QUẢ] Tổng hợp {0} | chi tiết {1} | lỗi {2}.' -f $summaryRows.Count, $detailRows.Count, $errorRows.Count)
    }
    if (-not $stopRequested -and $summaryRows.Count -eq 0 -and $errorRows.Count -gt 0) {
        Write-HddtLog ERROR '[KẾT QUẢ] Không lấy được hóa đơn nào; workbook chỉ chứa báo cáo lỗi. Mã thoát 2.'
        exit 2
    }
    exit 0
}
catch {
    if (Test-HddtStopRequested) {
        Write-HddtLog WARN ('[KẾT QUẢ] Đã dừng theo yêu cầu trước khi hoàn tất: {0}' -f $_.Exception.Message)
        exit 0
    }
    if ($loggingStarted) {
        Write-HddtLog ERROR $_.Exception.Message
        $savedLog = Get-HddtLogFile
        if (-not [string]::IsNullOrWhiteSpace($savedLog)) { Write-HddtLog ERROR ('Xem nhật ký: {0}' -f $savedLog) }
    }
    else {
        Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
    exit 1
}
