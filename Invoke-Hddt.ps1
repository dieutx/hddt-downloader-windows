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
. (Join-Path $PSScriptRoot 'src\Parallel.ps1')
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

# Áp kết quả tải/parse XML của một hóa đơn vào dòng tổng hợp đã dựng sẵn.
# Hóa đơn lỗi (trừ khi do dừng theo yêu cầu) sinh một dòng báo cáo lỗi.
function Add-HddtInvoiceResult {
    param($Binding, $Result)
    $summary = $Binding.Summary
    $invoice = $Binding.Invoice
    if ($Result.Ok) {
        $firstXml = $true
        foreach ($parsed in @($Result.Parsed)) {
            if ($firstXml) {
                Merge-HddtParsedSummary -Target $summary -Parsed $parsed.Summary
                $firstXml = $false
            }
            foreach ($row in @($parsed.Details)) { $Binding.DetailRows.Add($row) }
        }
        return
    }
    if ($Result.Stopped) { return }
    $Binding.ErrorRow = [pscustomobject]@{
        Direction = $invoice.Direction
        Source = $invoice.Source
        SellerTaxCode = $invoice.SellerTaxCode
        InvoiceTemplate = $invoice.InvoiceTemplate
        InvoiceSeries = $invoice.InvoiceSeries
        InvoiceNumber = $invoice.InvoiceNumber
        InvoiceDate = $invoice.InvoiceDate
        Stage = [string]$Result.Stage
        Endpoint = [string]$Result.Endpoint
        StatusCode = [int]$Result.StatusCode
        Attempts = [int]$Result.Attempts
        RetryAfterSeconds = [int]$Result.RetryAfterSeconds
        RecordedAt = [datetime]::Now
        Error = [string]$Result.Error
        FinalResult = 'Không tải được'
        Note = ''
    }
    Write-HddtLog WARN ('[TẢI XML] Không tải được {0}: {1}' -f (Get-InvoiceLabel -Invoice $invoice), $Result.Error)
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
    # Dựng trước toàn bộ dòng tổng hợp theo đúng thứ tự danh sách; hóa đơn tải
    # XML lỗi vẫn có dòng tổng hợp kèm dòng lỗi, nên không mất dữ liệu.
    foreach ($invoice in $allInvoices) {
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
        $summaryBindings.Add([pscustomobject]@{
            Summary = $summary
            Invoice = $invoice
            DetailRows = (New-Object System.Collections.Generic.List[object])
            ErrorRow = $null
        })
    }

    # Chỉ tải/parse XML mới chạy song song. Mỗi hóa đơn lấy từ chỉ số dùng chung
    # đúng một lần; kết quả trả về kèm Index nên khi gom lại vẫn giữ nguyên thứ
    # tự danh sách (không trùng, không thiếu dòng).
    $workerCount = [int]$config.DownloadWorkers
    if ($workerCount -gt 1 -and $summaryBindings.Count -gt 1) {
        Write-HddtLog INFO ('[TẢI XML] Chế độ tải: {0} luồng (tự điều tiết 1-{0} theo tải máy chủ).' -f $workerCount)
        $gate = New-HddtSharedGate -MaxWorkers $workerCount -BaseDelayMs $config.RequestDelayMs
        Set-HddtSharedGate $gate
        # Lưu ý PS 5.1: @() trên List[object] chứa pscustomobject có thể lỗi
        # "Argument types do not match", nên dùng ToArray() trực tiếp.
        $invoiceArray = $allInvoices.ToArray()
        $pool = Start-HddtXmlDownloadPool -Config $config -Invoices $invoiceArray -Gate $gate -Root $PSScriptRoot `
            -LogFile (Get-HddtLogFile) -LogLevel $config.LogLevel -LogToFile ([bool]$config.LogToFile) -WorkerCount $workerCount
        $applied = 0
        $totalCount = $summaryBindings.Count
        try {
            while ($true) {
                if (Test-HddtStopRequested) { break }
                $anyRunning = $false
                foreach ($worker in $pool.Workers) {
                    if (-not $worker.Handle.IsCompleted) { $anyRunning = $true; break }
                }
                $item = Get-HddtXmlPoolItem -Pool $pool
                while ($null -ne $item) {
                    $binding = $summaryBindings[[int]$item.Index]
                    Add-HddtInvoiceResult -Binding $binding -Result $item
                    $applied++
                    if (($applied % $config.ProgressEvery) -eq 0 -or $applied -eq $totalCount) {
                        $percent = if ($totalCount -eq 0) { 100 } else { [Math]::Floor(($applied * 100.0) / $totalCount) }
                        Write-HddtLog INFO ('[TẢI XML] [{0}/{1} | {2}%] {3} | XML {4} | +{5} dòng chi tiết' -f $applied, $totalCount, $percent, (Get-InvoiceLabel -Invoice $binding.Invoice), @($item.XmlFiles).Count, $binding.DetailRows.Count)
                    }
                    $item = Get-HddtXmlPoolItem -Pool $pool
                }
                if (-not $anyRunning -and $pool.Queue.Count -eq 0) { break }
                Start-Sleep -Milliseconds 50
            }
            $item = Get-HddtXmlPoolItem -Pool $pool
            while ($null -ne $item) {
                $binding = $summaryBindings[[int]$item.Index]
                Add-HddtInvoiceResult -Binding $binding -Result $item
                $applied++
                $item = Get-HddtXmlPoolItem -Pool $pool
            }
            if (Test-HddtStopRequested) {
                Write-HddtLog WARN ('[TẢI XML] Đã yêu cầu dừng; đã xử lý {0}/{1} hóa đơn.' -f $applied, $totalCount)
            }
        }
        finally {
            Complete-HddtXmlDownloadPool -Pool $pool
            Set-HddtSharedGate $null
        }
    }
    else {
        Write-HddtLog INFO '[TẢI XML] Chế độ tải: tuần tự.'
        $current = 0
        foreach ($binding in $summaryBindings) {
            if (Test-HddtStopRequested) {
                Write-HddtLog WARN ('[TẢI XML] Đã yêu cầu dừng; giữ lại {0} hóa đơn đã tải.' -f $current)
                break
            }
            $current++
            $result = Invoke-HddtInvoiceXmlJob -Config $config -Invoice $binding.Invoice
            Add-HddtInvoiceResult -Binding $binding -Result $result
            if (($current % $config.ProgressEvery) -eq 0 -or $current -eq $summaryBindings.Count) {
                $percent = if ($summaryBindings.Count -eq 0) { 100 } else { [Math]::Floor(($current * 100.0) / $summaryBindings.Count) }
                Write-HddtLog INFO ('[TẢI XML] [{0}/{1} | {2}%] {3} | XML {4} | +{5} dòng chi tiết' -f $current, $summaryBindings.Count, $percent, (Get-InvoiceLabel -Invoice $binding.Invoice), @($result.XmlFiles).Count, $binding.DetailRows.Count)
            }
        }
    }

    # Gom dòng chi tiết và dòng lỗi theo đúng thứ tự danh sách, bất kể luồng nào
    # hoàn thành trước.
    foreach ($binding in $summaryBindings) {
        foreach ($row in $binding.DetailRows) { $detailRows.Add($row) }
        if ($null -ne $binding.ErrorRow) { $errorRows.Add($binding.ErrorRow) }
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
        Write-HddtLog ERROR ('Chi tiết lỗi: ' + $_.ScriptStackTrace)
        $savedLog = Get-HddtLogFile
        if (-not [string]::IsNullOrWhiteSpace($savedLog)) { Write-HddtLog ERROR ('Xem nhật ký: {0}' -f $savedLog) }
    }
    else {
        Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
    exit 1
}
