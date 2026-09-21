[CmdletBinding()]
param([string]$EnvFile)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($EnvFile)) { $EnvFile = Join-Path $PSScriptRoot '.env' }

. (Join-Path $PSScriptRoot 'src\Logging.ps1')
. (Join-Path $PSScriptRoot 'src\Config.ps1')
. (Join-Path $PSScriptRoot 'src\Http.ps1')
. (Join-Path $PSScriptRoot 'src\InvoiceApi.ps1')
. (Join-Path $PSScriptRoot 'src\XmlParser.ps1')
. (Join-Path $PSScriptRoot 'src\ExcelExporter.ps1')

Initialize-HddtConsole
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$loggingStarted = $false

try {
    $config = Get-HddtConfig -EnvFile $EnvFile -RepositoryRoot $PSScriptRoot
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
    Write-HddtLog INFO ('Phạm vi: {0} | {1:dd/MM/yyyy} - {2:dd/MM/yyyy} | nguồn: {3}' -f ($config.Directions -join '+'), $config.FromDate, $config.ToDate, ($familyNames -join '+'))
    $adaptiveText = if ($config.AdaptiveThrottle) { 'bật' } else { 'tắt' }
    $xmlModeText = if ($config.RedownloadXml) { 'tải lại từ đầu' } else { 'tiếp tục/tái sử dụng' }
    Write-HddtLog INFO ('Mạng: tuần tự | page size {0} | giãn cách {1} ms | tự điều tiết {2} | retry {3} | timeout {4}s' -f $config.PageSize, $config.RequestDelayMs, $adaptiveText, $config.MaxRetries, $config.HttpTimeoutSeconds)
    Write-HddtLog INFO ('Chế độ XML: {0}.' -f $xmlModeText)
    Write-HddtLog INFO ('Đầu ra: {0}' -f $config.OutputWorkbook)
    if ($config.LogToFile) { Write-HddtLog INFO ('Nhật ký: {0}' -f $logFile) }
    Write-HddtLog INFO 'Token đã được nạp; request danh sách đầu tiên sẽ xác thực token.'

    $allInvoices = New-Object System.Collections.Generic.List[object]
    foreach ($direction in $config.Directions) {
        $beforeCount = $allInvoices.Count
        Write-HddtLog INFO ('Bắt đầu lấy danh sách {0}...' -f $direction)
        $items = @(Get-GdtInvoiceIndex -Config $config -Direction $direction)
        foreach ($item in $items) { $allInvoices.Add($item) }
        Write-HddtLog INFO ('Hoàn tất danh sách {0}: thêm {1} hóa đơn.' -f $direction, ($allInvoices.Count - $beforeCount))
    }

    Write-HddtLog INFO ('Tổng cộng {0} hóa đơn. Bắt đầu tải và parse XML.' -f $allInvoices.Count)
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $detailRows = New-Object System.Collections.Generic.List[object]
    $errorRows = New-Object System.Collections.Generic.List[object]
    $current = 0

    foreach ($invoice in $allInvoices) {
        $current++
        $label = Get-InvoiceLabel -Invoice $invoice
        $detailBefore = $detailRows.Count
        try {
            Write-HddtLog DEBUG ('Bắt đầu hóa đơn {0}/{1}: {2} [{3}]' -f $current, $allInvoices.Count, $label, $invoice.Source)
            $xmlFiles = @(Save-GdtInvoiceXml -Config $config -Invoice $invoice)
            foreach ($xmlFile in $xmlFiles) {
                $parsed = ConvertFrom-InvoiceXml -Path $xmlFile -Direction $invoice.Direction -Source $invoice.Source
                $summaryRows.Add($parsed.Summary)
                foreach ($row in $parsed.Details) { $detailRows.Add($row) }
            }

            if (($current % $config.ProgressEvery) -eq 0 -or $current -eq $allInvoices.Count) {
                $percent = if ($allInvoices.Count -eq 0) { 100 } else { [Math]::Floor(($current * 100.0) / $allInvoices.Count) }
                Write-HddtLog INFO ('[{0}/{1} | {2}%] OK {3} | XML {4} | dòng chi tiết +{5}' -f $current, $allInvoices.Count, $percent, $label, $xmlFiles.Count, ($detailRows.Count - $detailBefore))
            }
        }
        catch {
            $errorRows.Add([pscustomobject]@{ Direction=$invoice.Direction; Source=$invoice.Source; Invoice=$label; Error=$_.Exception.Message })
            Write-HddtLog WARN ('[{0}/{1}] Không tải được {2}: {3}' -f $current, $allInvoices.Count, $label, $_.Exception.Message)
        }
    }

    Write-HddtLog INFO ('Đang tạo workbook: tổng hợp {0}, chi tiết {1}, lỗi {2}.' -f $summaryRows.Count, $detailRows.Count, $errorRows.Count)
    Export-InvoiceWorkbook -Path $config.OutputWorkbook -SummaryRows ($summaryRows.ToArray()) -DetailRows ($detailRows.ToArray()) -ErrorRows ($errorRows.ToArray()) -Overwrite:$config.OverwriteOutput

    $stopwatch.Stop()
    Write-HddtLog INFO ('Hoàn tất sau {0:hh\:mm\:ss}. File: {1}' -f $stopwatch.Elapsed, $config.OutputWorkbook)
    Write-HddtLog INFO ('Kết quả: tổng hợp {0} | chi tiết {1} | lỗi {2}.' -f $summaryRows.Count, $detailRows.Count, $errorRows.Count)
    exit 0
}
catch {
    $stopwatch.Stop()
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
