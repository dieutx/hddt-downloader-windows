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
. (Join-Path $PSScriptRoot 'src\BrowserProfile.ps1')
. (Join-Path $PSScriptRoot 'src\XmlScheduler.ps1')

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

# Số hóa đơn đã nhận kết quả / tổng số, dùng cho thông báo lỗi trong log.
$script:HddtXmlInvoiceCount = 0
$script:HddtXmlProcessedCount = 0

# Dòng lỗi tải/parse; cùng cột với lỗi danh sách để xuất chung sheet.
function New-HddtDownloadErrorRow {
    param(
        [Parameter(Mandatory = $true)]$Invoice,
        [Parameter(Mandatory = $true)][string]$Stage,
        [string]$Endpoint = '',
        [int]$StatusCode = 0,
        [int]$Attempts = 0,
        [int]$RetryAfterSeconds = 0,
        [Parameter(Mandatory = $true)][string]$ErrorText,
        [string]$FinalResult = 'Không tải được'
    )
    return [pscustomobject]@{
        Direction = $Invoice.Direction
        Source = $Invoice.Source
        SellerTaxCode = $Invoice.SellerTaxCode
        InvoiceTemplate = $Invoice.InvoiceTemplate
        InvoiceSeries = $Invoice.InvoiceSeries
        InvoiceNumber = $Invoice.InvoiceNumber
        InvoiceDate = $Invoice.InvoiceDate
        Stage = $Stage
        Endpoint = $Endpoint
        StatusCode = $StatusCode
        Attempts = $Attempts
        RetryAfterSeconds = $RetryAfterSeconds
        RecordedAt = [datetime]::Now
        Error = $ErrorText
        FinalResult = $FinalResult
        Note = ''
    }
}

# Ghi một kết quả tải từ worker vào đúng dòng tổng hợp theo chỉ số hóa đơn:
# thành công thì parse XML và cộng dòng chi tiết, thất bại thì giữ sẵn một
# dòng lỗi. Kết quả về theo thứ tự bất kỳ nhưng workbook vẫn đúng thứ tự cũ.
function Add-HddtXmlResultRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Binding,
        [Parameter(Mandatory = $true)]$Result
    )

    $Binding.HasResult = $true
    if (-not $Result.Success) {
        # Khi người dùng dừng thì lỗi "đã dừng" không phải lỗi tải; giữ nguyên
        # hành vi cũ là không ghi dòng lỗi cho request bị hủy giữa chừng.
        if (-not (Test-HddtStopRequested)) {
            $Binding.DownloadError = New-HddtDownloadErrorRow -Invoice $Binding.Invoice `
                -Stage $Result.Stage -Endpoint $Result.Endpoint -StatusCode $Result.StatusCode `
                -Attempts $Result.Attempts -RetryAfterSeconds $Result.RetryAfterSeconds `
                -ErrorText $Result.ErrorMessage
            Write-HddtLog WARN ('[TẢI XML] Không tải được {0} ({1}/{2}): {3}' -f $Result.Label, $script:HddtXmlProcessedCount, $script:HddtXmlInvoiceCount, $Result.ErrorMessage)
        }
        return 0
    }

    $detailAdded = 0
    $stage = 'Tải XML'
    try {
        $firstXml = $true
        foreach ($xmlFile in @($Result.XmlFiles)) {
            $stage = 'Parse XML'
            $parsed = ConvertFrom-InvoiceXml -Path $xmlFile -Direction $Binding.Invoice.Direction -Source $Binding.Invoice.Source
            if ($firstXml) {
                Merge-HddtParsedSummary -Target $Binding.Summary -Parsed $parsed.Summary
                $firstXml = $false
            }
            foreach ($row in $parsed.Details) {
                $Binding.Details.Add($row)
                $detailAdded++
            }
        }
    }
    catch {
        $Binding.DownloadError = New-HddtDownloadErrorRow -Invoice $Binding.Invoice `
            -Stage $stage -ErrorText $_.Exception.Message
        Write-HddtLog WARN ('[TẢI XML] Không tải được {0} ({1}/{2}): {3}' -f (Get-InvoiceLabel -Invoice $Binding.Invoice), $script:HddtXmlProcessedCount, $script:HddtXmlInvoiceCount, $_.Exception.Message)
    }
    return $detailAdded
}

# Nhận một kết quả từ pipeline: ghi vào đúng dòng theo chỉ số rồi in tiến độ.
function Add-HddtXmlPipelineResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$SummaryBindings,
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)][int]$ProgressEvery
    )

    $script:HddtXmlProcessedCount++
    $binding = $SummaryBindings[[int]$Result.PipelineIndex]
    $detailAdded = Add-HddtXmlResultRows -Binding $binding -Result $Result
    $processed = $script:HddtXmlProcessedCount
    $total = $script:HddtXmlInvoiceCount
    if ($total -gt 0 -and (($processed % $ProgressEvery) -eq 0 -or $processed -eq $total)) {
        $percent = [Math]::Floor(($processed * 100.0) / $total)
        # Hiển thị worker nào vừa xử lý và còn bao nhiêu worker đang chạy để
        # theo dõi mức song song thực tế, không chỉ tiến độ hóa đơn.
        $resultWorker = 0
        $workerProperty = $Result.PSObject.Properties['WorkerIndex']
        if ($null -ne $workerProperty) { $resultWorker = [int]$workerProperty.Value }
        $workerTotal = [int](Get-HddtSharedValue -Key 'WorkerCount' -Default 0)
        $activeWorkers = 0
        $snapshot = Get-GdtXmlThrottleSnapshot
        if ($null -ne $snapshot) { $activeWorkers = [int]$snapshot.ActiveCount }
        Write-HddtLog INFO ('[TẢI XML] [{0}/{1} | {2}%] Worker {3}/{4} | đang chạy {5}/{4} | {6} | XML {7} | +{8} dòng chi tiết' -f $processed, $total, $percent, $resultWorker, $workerTotal, $activeWorkers, $Result.Label, @($Result.XmlFiles).Count, $detailAdded)
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
    Write-HddtLog INFO ('[CẤU HÌNH] Tải XML: tối đa {0} kết nối (khởi đầu {1}) | giãn cách {2} ms/request.' -f $config.XmlMaxConcurrency, $config.XmlConcurrency, $config.XmlRequestIntervalMs)
    Write-GdtBrowserProfileLog -Config $config
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
    $summaryBindings = New-Object System.Collections.Generic.List[object]
    $detailRows = New-Object System.Collections.Generic.List[object]
    $errorRows = New-Object System.Collections.Generic.List[object]
    foreach ($indexError in $indexErrors) { $errorRows.Add($indexError) }

    # Dựng sẵn một binding cho từng hóa đơn đúng theo thứ tự danh sách: worker
    # trả kết quả về thứ tự bất kỳ nhưng workbook vẫn đúng thứ tự cũ.
    $invoicesByIndex = $allInvoices.ToArray()
    for ($invoiceIndex = 0; $invoiceIndex -lt $invoicesByIndex.Count; $invoiceIndex++) {
        $invoice = $invoicesByIndex[$invoiceIndex]
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
        $summaryBindings.Add([pscustomobject]@{
            Summary = $summary
            Invoice = $invoice
            Index = $invoiceIndex
            Details = (New-Object System.Collections.Generic.List[object])
            DownloadError = $null
            HasResult = $false
        })
    }

    $script:HddtXmlInvoiceCount = $invoicesByIndex.Count
    $script:HddtXmlProcessedCount = 0

    # Windows PowerShell 5.1 mặc định chỉ 2 kết nối tới cùng một host; nâng
    # đủ cho số worker để các request không xếp hàng chờ nhau.
    try {
        $currentConnectionLimit = [Net.ServicePointManager]::DefaultConnectionLimit
        $desiredConnectionLimit = $config.XmlMaxConcurrency + 1
        if ($desiredConnectionLimit -gt $currentConnectionLimit) {
            [Net.ServicePointManager]::DefaultConnectionLimit = $desiredConnectionLimit
        }
    }
    catch { }

    $sharedState = $null
    $pipeline = $null
    if ($invoicesByIndex.Count -gt 0) {
        Write-HddtLog INFO ('[TẢI XML] Bắt đầu tải song song: tối đa {0} kết nối | giãn cách {1} ms | khởi đầu {2} kết nối.' -f $config.XmlMaxConcurrency, $config.XmlRequestIntervalMs, $config.XmlConcurrency)
        $sharedState = New-HddtXmlSharedState -Config $config
        try {
            $pipeline = Start-HddtXmlPipeline -Config $config -Invoices $invoicesByIndex -Root $PSScriptRoot
            while ($true) {
                foreach ($result in @(Receive-HddtXmlPipeline -Pipeline $pipeline)) {
                    Add-HddtXmlPipelineResult -SummaryBindings $summaryBindings -Result $result -ProgressEvery $config.ProgressEvery
                }
                if (Test-HddtXmlPipelineCompleted -Pipeline $pipeline) {
                    # Worker có thể vừa thêm kết quả ngay sau lần nhận trước.
                    foreach ($result in @(Receive-HddtXmlPipeline -Pipeline $pipeline)) {
                        Add-HddtXmlPipelineResult -SummaryBindings $summaryBindings -Result $result -ProgressEvery $config.ProgressEvery
                    }
                    break
                }
                if (Test-HddtStopRequested) { break }
                Start-Sleep -Milliseconds 200
            }
        }
        finally {
            # Dù dừng hay chạy xong cũng phải chờ worker thoát và lấy nốt log,
            # kết quả còn trong hàng đợi trước khi giải phóng runspace.
            $stopResults = @()
            if ($null -ne $pipeline) { $stopResults = @(Stop-HddtXmlPipeline -Pipeline $pipeline) }
            foreach ($result in $stopResults) {
                Add-HddtXmlPipelineResult -SummaryBindings $summaryBindings -Result $result -ProgressEvery $config.ProgressEvery
            }
        }
    }

    if ($null -ne $sharedState) {
        $metrics = Complete-HddtSharedState -Config $config
        if ($null -ne $metrics) {
            Write-HddtLog INFO ('[TẢI XML] Thống kê: worker {0} | {1} request thành công | kết nối cao nhất {2} | rate-limit {3} | trung bình {4} ms/request | làm mới token {5} | chạy {6}.' -f $metrics.WorkerCount, $metrics.CompletedCount, $metrics.PeakConcurrency, $metrics.RateLimitCount, $metrics.AverageResponseTimeMs, $metrics.AuthRefreshCount, (Get-HddtElapsedText))
        }
    }

    if (Test-HddtStopRequested -and $invoicesByIndex.Count -gt 0) {
        $keptCount = 0
        foreach ($binding in $summaryBindings) { if ($binding.HasResult) { $keptCount++ } }
        Write-HddtLog WARN ('[TẢI XML] Đã yêu cầu dừng; giữ lại {0} hóa đơn đã tải.' -f $keptCount)
    }

    # Worker kết thúc mà chưa trả kết quả (lỗi không mong muốn): khi chưa dừng
    # thì mọi hóa đơn vẫn phải có một dòng lỗi thay vì mất lặng khỏi workbook.
    if (-not (Test-HddtStopRequested)) {
        foreach ($binding in $summaryBindings) {
            if ($binding.HasResult) { continue }
            $binding.HasResult = $true
            $binding.DownloadError = New-HddtDownloadErrorRow -Invoice $binding.Invoice -Stage 'Tải XML' `
                -ErrorText 'Không nhận được kết quả tải cho hóa đơn này; worker kết thúc trước khi xử lý.'
            Write-HddtLog WARN ('[TẢI XML] Không tải được {0} ({1}/{2}): không nhận được kết quả từ worker.' -f (Get-InvoiceLabel -Invoice $binding.Invoice), $script:HddtXmlProcessedCount, $script:HddtXmlInvoiceCount)
        }
    }

    # Đóng gói theo đúng thứ tự chỉ số hóa đơn; kết quả về không theo thứ tự.
    $summaryRows = New-Object System.Collections.Generic.List[object]
    foreach ($binding in $summaryBindings) {
        if (-not $binding.HasResult) { continue }
        $summaryRows.Add($binding.Summary)
        foreach ($row in $binding.Details) { $detailRows.Add($row) }
        if ($null -ne $binding.DownloadError) { $errorRows.Add($binding.DownloadError) }
    }

    # Chuỗi hóa đơn liên quan và thông tin liên quan (tương ứng
    # ProcessRelatedInvoiceApis trong VBA). Ghi thẳng vào dòng tổng hợp.
    $relatedColumns = @('RelatedChain', 'OriginalInvoiceType', 'OriginalTemplateCode', 'OriginalSeries', 'OriginalNumber', 'OriginalDate', 'OriginalNote', 'RelatedInfo')
    # Chỉ hóa đơn đã có kết quả (dòng tổng hợp sẽ được xuất) mới cần gọi API
    # liên quan; hóa đơn chưa xử lý khi dừng không tốn request và không ghi.
    $relatedCandidates = @($summaryBindings | Where-Object { $_.HasResult -and (Test-GdtRelatedInvoice -Invoice $_.Invoice) })
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
