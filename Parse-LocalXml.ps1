[CmdletBinding()]
param([string]$EnvFile)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($EnvFile)) {
    $EnvFile = Join-Path $PSScriptRoot '.env'
}

. (Join-Path $PSScriptRoot 'src\Logging.ps1')
. (Join-Path $PSScriptRoot 'src\Config.ps1')
. (Join-Path $PSScriptRoot 'src\XmlParser.ps1')
. (Join-Path $PSScriptRoot 'src\ExcelExporter.ps1')

Initialize-HddtConsole
$loggingStarted = $false
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

try {
    $values = Read-DotEnvFile -Path $EnvFile
    $sourceDirectoryValue = Get-RequiredEnvValue -Values $values -Name 'LOCAL_XML_DIR'
    $sourceDirectory = Resolve-RepositoryPath -RepositoryRoot $PSScriptRoot -Value $sourceDirectoryValue
    if (-not (Test-Path -LiteralPath $sourceDirectory -PathType Container)) {
        throw "Không tìm thấy thư mục XML: $sourceDirectory"
    }

    $direction = (Get-EnvValue -Values $values -Name 'LOCAL_DIRECTION' -Default 'auto').ToLowerInvariant()
    if ($direction -notin @('auto', 'purchase', 'sold')) {
        throw 'LOCAL_DIRECTION phải là auto, purchase hoặc sold.'
    }

    $outputDirectory = Resolve-RepositoryPath -RepositoryRoot $PSScriptRoot -Value (Get-EnvValue -Values $values -Name 'OUTPUT_DIR' -Default 'output')
    $outputName = Get-EnvValue -Values $values -Name 'LOCAL_OUTPUT_XLSX' -Default 'HoaDonDienTu_Local.xlsx'
    if ([IO.Path]::GetExtension($outputName).ToLowerInvariant() -ne '.xlsx') {
        throw 'LOCAL_OUTPUT_XLSX phải có phần mở rộng .xlsx.'
    }
    $outputWorkbook = Join-Path $outputDirectory $outputName
    $overwrite = ConvertTo-EnvBoolean -Name 'OVERWRITE_OUTPUT' -Value (Get-EnvValue -Values $values -Name 'OVERWRITE_OUTPUT' -Default 'false')
    $progressEvery = [int](Get-EnvValue -Values $values -Name 'PROGRESS_EVERY' -Default '1')
    $logLevel = (Get-EnvValue -Values $values -Name 'LOG_LEVEL' -Default 'info').ToLowerInvariant()
    $logToFile = ConvertTo-EnvBoolean -Name 'LOG_TO_FILE' -Value (Get-EnvValue -Values $values -Name 'LOG_TO_FILE' -Default 'true')
    if ($progressEvery -lt 1 -or $progressEvery -gt 1000) { throw 'PROGRESS_EVERY phải nằm trong khoảng 1-1000.' }
    if ($logLevel -notin @('debug','info','warn','error')) { throw 'LOG_LEVEL không hợp lệ.' }

    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
    $logFile = Start-HddtLogging -Directory (Join-Path $outputDirectory 'logs') -Level $logLevel -ToFile $logToFile -RunName 'parse-local'
    $loggingStarted = $true
    Write-HddtLog INFO ('[PARSE XML] Bắt đầu: {0}' -f $sourceDirectory)
    $directionText = if ($direction -eq 'auto') { 'tự nhận diện theo thư mục/tên file' } else { $direction }
    Write-HddtLog INFO ('[PARSE XML] Loại: {0} | đầu ra: {1}' -f $directionText, $outputWorkbook)
    if ($logToFile) { Write-HddtLog INFO ('[PARSE XML] Nhật ký: {0}' -f $logFile) }

    $xmlFiles = @(Get-ChildItem -LiteralPath $sourceDirectory -Recurse -File -Filter '*.xml' | Sort-Object FullName)
    if ($xmlFiles.Count -eq 0) { throw "Thư mục không có file XML: $sourceDirectory" }
    Write-HddtLog INFO ('[PARSE XML] Tìm thấy {0} file XML.' -f $xmlFiles.Count)

    $summaryRows = New-Object System.Collections.Generic.List[object]
    $detailRows = New-Object System.Collections.Generic.List[object]
    $errorRows = New-Object System.Collections.Generic.List[object]
    $current = 0

    foreach ($file in $xmlFiles) {
        $current++
        $fileDirection = $direction
        Write-HddtLog DEBUG ('Đang parse {0}' -f $file.FullName)
        try {
            $fileDirection = Resolve-LocalInvoiceDirection -Path $file.FullName -RootDirectory $sourceDirectory -ConfiguredDirection $direction
            $parsed = ConvertFrom-InvoiceXml -Path $file.FullName -Direction $fileDirection -Source 'local-xml'
            $summaryRows.Add($parsed.Summary)
            foreach ($detail in $parsed.Details) { $detailRows.Add($detail) }
            if (($current % $progressEvery) -eq 0 -or $current -eq $xmlFiles.Count) {
                $percent = [Math]::Floor(($current * 100.0) / $xmlFiles.Count)
                Write-HddtLog INFO ('[PARSE XML] [{0}/{1} | {2}%] {3} | +{4} dòng chi tiết' -f $current, $xmlFiles.Count, $percent, $file.Name, @($parsed.Details).Count)
            }
        }
        catch {
            $errorRows.Add([pscustomobject]@{
                Direction = $fileDirection
                Source = 'local-xml'
                Invoice = $file.FullName
                Error = $_.Exception.Message
            })
            Write-HddtLog WARN ('[PARSE XML] Không parse được {2} ({0}/{1}): {3}' -f $current, $xmlFiles.Count, $file.Name, $_.Exception.Message)
        }
    }

    Write-HddtLog INFO ('[XUẤT FILE] Tạo workbook: tổng hợp {0} | chi tiết {1} | lỗi {2}.' -f $summaryRows.Count, $detailRows.Count, $errorRows.Count)
    Export-InvoiceWorkbook -Path $outputWorkbook -SummaryRows ($summaryRows.ToArray()) -DetailRows ($detailRows.ToArray()) -ErrorRows ($errorRows.ToArray()) -Overwrite:$overwrite
    $stopwatch.Stop()
    Write-HddtLog INFO ('[KẾT QUẢ] Hoàn tất | chạy {0:hh\:mm\:ss} | file: {1}' -f $stopwatch.Elapsed, $outputWorkbook)
    Write-HddtLog INFO ('[KẾT QUẢ] XML {0} | tổng hợp {1} | chi tiết {2} | lỗi {3}.' -f $xmlFiles.Count, $summaryRows.Count, $detailRows.Count, $errorRows.Count)
    exit 0
}
catch {
    $stopwatch.Stop()
    if ($loggingStarted) { Write-HddtLog ERROR $_.Exception.Message }
    else { Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red }
    exit 1
}
