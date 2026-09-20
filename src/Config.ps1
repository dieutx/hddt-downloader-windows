Set-StrictMode -Version 2.0

function Read-DotEnvFile {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Không tìm thấy file cấu hình: $Path. Hãy sao chép .env.example thành .env rồi điền thông tin."
    }

    $values = @{}
    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = $rawLine.Trim()
        if ($line.Length -eq 0 -or $line.StartsWith('#')) { continue }

        $separator = $line.IndexOf('=')
        if ($separator -lt 1) {
            throw "Dòng .env không hợp lệ: $rawLine"
        }

        $name = $line.Substring(0, $separator).Trim()
        $value = $line.Substring($separator + 1).Trim()
        if ($value.Length -ge 2) {
            if (($value.StartsWith('"') -and $value.EndsWith('"')) -or ($value.StartsWith("'") -and $value.EndsWith("'"))) {
                $value = $value.Substring(1, $value.Length - 2)
            }
        }
        if ($name -notmatch '^[A-Z][A-Z0-9_]*$') {
            throw "Tên biến .env không hợp lệ: $name"
        }
        $values[$name] = $value
    }
    return $values
}

function Get-RequiredEnvValue {
    param([hashtable]$Values, [string]$Name)
    if (-not $Values.ContainsKey($Name) -or [string]::IsNullOrWhiteSpace($Values[$Name])) {
        throw "Thiếu giá trị bắt buộc trong .env: $Name"
    }
    return [string]$Values[$Name]
}

function Get-EnvValue {
    param([hashtable]$Values, [string]$Name, [string]$Default)
    if ($Values.ContainsKey($Name) -and -not [string]::IsNullOrWhiteSpace($Values[$Name])) {
        return [string]$Values[$Name]
    }
    return $Default
}

function ConvertTo-EnvBoolean {
    param([string]$Name, [string]$Value)
    switch ($Value.Trim().ToLowerInvariant()) {
        'true' { return $true }
        '1' { return $true }
        'yes' { return $true }
        'false' { return $false }
        '0' { return $false }
        'no' { return $false }
        default { throw "$Name phải là true hoặc false." }
    }
}

function Resolve-RepositoryPath {
    param([string]$RepositoryRoot, [string]$Value)
    if ([System.IO.Path]::IsPathRooted($Value)) {
        return [System.IO.Path]::GetFullPath($Value)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $RepositoryRoot $Value))
}

function Get-HddtConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EnvFile,
        [Parameter(Mandatory = $true)][string]$RepositoryRoot
    )

    $values = Read-DotEnvFile -Path $EnvFile
    $token = (Get-RequiredEnvValue -Values $values -Name 'GDT_TOKEN').Trim()
    $token = $token -replace '^(?i)Bearer\s+', ''
    if ($token.IndexOf("`r") -ge 0 -or $token.IndexOf("`n") -ge 0 -or $token.Length -lt 20) {
        throw 'GDT_TOKEN không hợp lệ.'
    }

    $direction = (Get-EnvValue -Values $values -Name 'INVOICE_DIRECTION' -Default 'both').ToLowerInvariant()
    switch ($direction) {
        'purchase' { $directions = @('purchase') }
        'sold' { $directions = @('sold') }
        'both' { $directions = @('purchase', 'sold') }
        default { throw 'INVOICE_DIRECTION phải là purchase, sold hoặc both.' }
    }

    try {
        $fromDate = [datetime]::ParseExact((Get-RequiredEnvValue $values 'FROM_DATE'), 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
        $toDate = [datetime]::ParseExact((Get-RequiredEnvValue $values 'TO_DATE'), 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw 'FROM_DATE và TO_DATE phải đúng định dạng dd/MM/yyyy.'
    }
    if ($toDate -lt $fromDate) { throw 'TO_DATE không được nhỏ hơn FROM_DATE.' }

    $pageSize = [int](Get-EnvValue $values 'PAGE_SIZE' '50')
    $delay = [int](Get-EnvValue $values 'REQUEST_DELAY_MS' '600')
    $retries = [int](Get-EnvValue $values 'MAX_RETRIES' '4')
    $timeout = [int](Get-EnvValue $values 'HTTP_TIMEOUT_SECONDS' '90')
    $progressEvery = [int](Get-EnvValue $values 'PROGRESS_EVERY' '1')
    $logLevel = (Get-EnvValue $values 'LOG_LEVEL' 'info').ToLowerInvariant()
    if ($pageSize -lt 1 -or $pageSize -gt 100) { throw 'PAGE_SIZE phải nằm trong khoảng 1-100.' }
    if ($delay -lt 0 -or $delay -gt 60000) { throw 'REQUEST_DELAY_MS phải nằm trong khoảng 0-60000.' }
    if ($retries -lt 0 -or $retries -gt 10) { throw 'MAX_RETRIES phải nằm trong khoảng 0-10.' }
    if ($timeout -lt 5 -or $timeout -gt 600) { throw 'HTTP_TIMEOUT_SECONDS phải nằm trong khoảng 5-600.' }
    if ($progressEvery -lt 1 -or $progressEvery -gt 1000) { throw 'PROGRESS_EVERY phải nằm trong khoảng 1-1000.' }
    if ($logLevel -notin @('debug', 'info', 'warn', 'error')) { throw 'LOG_LEVEL phải là debug, info, warn hoặc error.' }

    $includeRegular = ConvertTo-EnvBoolean 'INCLUDE_REGULAR' (Get-EnvValue $values 'INCLUDE_REGULAR' 'true')
    $includeSco = ConvertTo-EnvBoolean 'INCLUDE_SCO' (Get-EnvValue $values 'INCLUDE_SCO' 'true')
    if (-not $includeRegular -and -not $includeSco) { throw 'Phải bật ít nhất một trong INCLUDE_REGULAR hoặc INCLUDE_SCO.' }

    $outputDirectory = Resolve-RepositoryPath $RepositoryRoot (Get-EnvValue $values 'OUTPUT_DIR' 'output')
    $outputName = Get-EnvValue $values 'OUTPUT_XLSX' 'HoaDonDienTu.xlsx'
    if ([System.IO.Path]::GetExtension($outputName).ToLowerInvariant() -ne '.xlsx') {
        throw 'OUTPUT_XLSX phải có phần mở rộng .xlsx.'
    }

    return [pscustomobject]@{
        Token = $token
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        Directions = $directions
        FromDate = $fromDate
        ToDate = $toDate
        PageSize = $pageSize
        RequestDelayMs = $delay
        MaxRetries = $retries
        HttpTimeoutSeconds = $timeout
        ProgressEvery = $progressEvery
        LogLevel = $logLevel
        LogToFile = ConvertTo-EnvBoolean 'LOG_TO_FILE' (Get-EnvValue $values 'LOG_TO_FILE' 'true')
        AdaptiveThrottle = ConvertTo-EnvBoolean 'ADAPTIVE_THROTTLE' (Get-EnvValue $values 'ADAPTIVE_THROTTLE' 'true')
        IncludeRegular = $includeRegular
        IncludeSco = $includeSco
        RedownloadXml = ConvertTo-EnvBoolean 'REDOWNLOAD_XML' (Get-EnvValue $values 'REDOWNLOAD_XML' 'false')
        OverwriteOutput = ConvertTo-EnvBoolean 'OVERWRITE_OUTPUT' (Get-EnvValue $values 'OVERWRITE_OUTPUT' 'false')
        OutputDirectory = $outputDirectory
        OutputWorkbook = Join-Path $outputDirectory $outputName
        XmlDirectory = Join-Path $outputDirectory 'xml'
        LogDirectory = Join-Path $outputDirectory 'logs'
    }
}
