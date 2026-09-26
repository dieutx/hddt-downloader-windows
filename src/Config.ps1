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
            else {
                # Cho phép inline comment kiểu `.env.example` (sau khoảng trắng),
                # nhưng không đụng vào dấu # nằm trong mật khẩu.
                $commentIndex = $value.IndexOf(' ;', [StringComparison]::Ordinal)
                if ($commentIndex -lt 0) { $commentIndex = $value.IndexOf("`t#", [StringComparison]::Ordinal) }
                if ($commentIndex -ge 0) { $value = $value.Substring(0, $commentIndex).Trim() }
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

# Đọc một trường của object cấu hình một cách an toàn với Set-StrictMode:
# object cấu hình do test tạo có thể thiếu trường mới, và đọc property không
# tồn tại sẽ ném lỗi thay vì trả về $Default.
function Get-HddtConfigValue {
    param($Config, [string]$Name, $Default = $null)
    if ($null -eq $Config) { return $Default }
    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $Default }
    return $property.Value
}

function Resolve-RepositoryPath {
    param([string]$RepositoryRoot, [string]$Value)
    if ([System.IO.Path]::IsPathRooted($Value)) {
        return [System.IO.Path]::GetFullPath($Value)
    }
    return [System.IO.Path]::GetFullPath((Join-Path $RepositoryRoot $Value))
}

function Read-HddtInteractiveValue {
    param([string]$Existing, [string]$Prompt, [string]$Default = '')
    $effectiveDefault = if ([string]::IsNullOrWhiteSpace($Existing)) { $Default } else { $Existing }
    $suffix = if ([string]::IsNullOrWhiteSpace($effectiveDefault)) { '' } else { " [$effectiveDefault]" }
    $answer = Read-Host ($Prompt + $suffix)
    if ([string]::IsNullOrWhiteSpace([string]$answer)) { return $effectiveDefault }
    return ([string]$answer).Trim()
}

function Read-HddtInteractiveSecret {
    param([string]$Existing, [string]$Prompt)
    $answer = Read-Host ($Prompt + ' (Enter để giữ giá trị hiện có; CLEAR để xóa)') -AsSecureString
    if ($null -eq $answer) { return $Existing }
    if ($answer -isnot [Security.SecureString]) {
        $plain = [string]$answer
    }
    else {
        $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($answer)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
    }
    if ([string]::IsNullOrWhiteSpace($plain)) { return $Existing }
    if ($plain.Trim() -match '^(?i:CLEAR|__CLEAR__)$') { return '' }
    return $plain
}

function Read-HddtInteractiveDate {
    param([string]$Existing, [string]$Prompt, [datetime]$Default)
    $defaultText = if ([string]::IsNullOrWhiteSpace($Existing)) { $Default.ToString('dd/MM/yyyy') } else { $Existing }
    while ($true) {
        $answer = Read-HddtInteractiveValue -Existing $Existing -Prompt $Prompt -Default $defaultText
        $date = [datetime]::MinValue
        if ([datetime]::TryParseExact($answer, 'dd/MM/yyyy', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)) {
            return $date.ToString('dd/MM/yyyy')
        }
        Write-Host 'Ngày không hợp lệ; nhập theo định dạng dd/MM/yyyy.' -ForegroundColor Yellow
    }
}

function Read-HddtInteractiveBoolean {
    param([string]$Existing, [string]$Prompt, [bool]$Default)
    $defaultText = if ($Default) { 'true' } else { 'false' }
    while ($true) {
        $answer = Read-HddtInteractiveValue -Existing $Existing -Prompt ($Prompt + ' (true/false)') -Default $defaultText
        try { return (ConvertTo-EnvBoolean $Prompt $answer) }
        catch { Write-Host 'Chỉ nhập true hoặc false.' -ForegroundColor Yellow }
    }
}

function Copy-HddtConfigValues {
    param([hashtable]$Values)
    $copy = @{}
    if ($null -ne $Values) {
        foreach ($key in $Values.Keys) { $copy[[string]$key] = [string]$Values[$key] }
    }
    return $copy
}

function Complete-HddtInteractiveValues {
    param([hashtable]$Values)
    $valuesCopy = Copy-HddtConfigValues $Values
    $today = Get-Date
    $defaultFrom = $today.Date.AddDays(1 - $today.Day)
    $defaultTo = $today.Date

    $valuesCopy['GDT_USERNAME'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'GDT_USERNAME' '') -Prompt 'GDT_USERNAME' -Default ''
    $valuesCopy['GDT_PASSWORD'] = Read-HddtInteractiveSecret -Existing (Get-EnvValue $valuesCopy 'GDT_PASSWORD' '') -Prompt 'GDT_PASSWORD'

    $valuesCopy['INVOICE_DIRECTION'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'INVOICE_DIRECTION' 'both') -Prompt 'INVOICE_DIRECTION (purchase/sold/both)' -Default 'both'
    $valuesCopy['FROM_DATE'] = Read-HddtInteractiveDate -Existing (Get-EnvValue $valuesCopy 'FROM_DATE' '') -Prompt 'FROM_DATE' -Default $defaultFrom
    $valuesCopy['TO_DATE'] = Read-HddtInteractiveDate -Existing (Get-EnvValue $valuesCopy 'TO_DATE' '') -Prompt 'TO_DATE' -Default $defaultTo
    $valuesCopy['OUTPUT_DIR'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'OUTPUT_DIR' 'output') -Prompt 'OUTPUT_DIR' -Default 'output'
    $valuesCopy['OUTPUT_XLSX'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'OUTPUT_XLSX' 'HoaDonDienTu.xlsx') -Prompt 'OUTPUT_XLSX' -Default 'HoaDonDienTu.xlsx'

    $valuesCopy['INCLUDE_REGULAR'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'INCLUDE_REGULAR' 'true') -Prompt 'INCLUDE_REGULAR' -Default $true
    $valuesCopy['INCLUDE_SCO'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'INCLUDE_SCO' 'true') -Prompt 'INCLUDE_SCO' -Default $true
    $valuesCopy['FETCH_RELATED'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'FETCH_RELATED' 'true') -Prompt 'FETCH_RELATED' -Default $true
    $valuesCopy['REDOWNLOAD_XML'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'REDOWNLOAD_XML' 'false') -Prompt 'REDOWNLOAD_XML' -Default $false
    $valuesCopy['OVERWRITE_OUTPUT'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'OVERWRITE_OUTPUT' 'false') -Prompt 'OVERWRITE_OUTPUT' -Default $false
    $proxyUrl = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'PROXY_URL' '') -Prompt 'PROXY_URL (http://host:port; CLEAR để tắt proxy)' -Default ''
    if ($proxyUrl.Trim() -match '^(?i:CLEAR|__CLEAR__)$') { $proxyUrl = '' }
    $valuesCopy['PROXY_URL'] = $proxyUrl
    if ([string]::IsNullOrWhiteSpace($proxyUrl)) {
        $valuesCopy['PROXY_USERNAME'] = ''
        $valuesCopy['PROXY_PASSWORD'] = ''
    }
    else {
        $valuesCopy['PROXY_USERNAME'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'PROXY_USERNAME' '') -Prompt 'PROXY_USERNAME (để trống nếu proxy không cần đăng nhập)' -Default ''
        $valuesCopy['PROXY_PASSWORD'] = Read-HddtInteractiveSecret -Existing (Get-EnvValue $valuesCopy 'PROXY_PASSWORD' '') -Prompt 'PROXY_PASSWORD'
    }
    return $valuesCopy
}

function Complete-HddtLocalInteractiveValues {
    param([hashtable]$Values)
    $valuesCopy = Copy-HddtConfigValues $Values
    $valuesCopy['LOCAL_XML_DIR'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'LOCAL_XML_DIR' '') -Prompt 'LOCAL_XML_DIR' -Default ''
    $valuesCopy['LOCAL_DIRECTION'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'LOCAL_DIRECTION' 'auto') -Prompt 'LOCAL_DIRECTION (auto/purchase/sold)' -Default 'auto'
    $valuesCopy['OUTPUT_DIR'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'OUTPUT_DIR' 'output') -Prompt 'OUTPUT_DIR' -Default 'output'
    $valuesCopy['LOCAL_OUTPUT_XLSX'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'LOCAL_OUTPUT_XLSX' 'HoaDonDienTu_Local.xlsx') -Prompt 'LOCAL_OUTPUT_XLSX' -Default 'HoaDonDienTu_Local.xlsx'
    $valuesCopy['OVERWRITE_OUTPUT'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'OVERWRITE_OUTPUT' 'false') -Prompt 'OVERWRITE_OUTPUT' -Default $false
    $valuesCopy['PROGRESS_EVERY'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'PROGRESS_EVERY' '1') -Prompt 'PROGRESS_EVERY' -Default '1'
    $valuesCopy['LOG_LEVEL'] = Read-HddtInteractiveValue -Existing (Get-EnvValue $valuesCopy 'LOG_LEVEL' 'info') -Prompt 'LOG_LEVEL' -Default 'info'
    $valuesCopy['LOG_TO_FILE'] = Read-HddtInteractiveBoolean -Existing (Get-EnvValue $valuesCopy 'LOG_TO_FILE' 'true') -Prompt 'LOG_TO_FILE' -Default $true
    return $valuesCopy
}

function Get-HddtConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$EnvFile,
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [switch]$Interactive
    )

    if ($Interactive) {
        $existingValues = @{}
        if (Test-Path -LiteralPath $EnvFile -PathType Leaf) { $existingValues = Read-DotEnvFile -Path $EnvFile }
        $values = Complete-HddtInteractiveValues -Values $existingValues
    }
    else {
        $values = Read-DotEnvFile -Path $EnvFile
    }

    # Chỉ dùng đăng nhập bằng tài khoản GDT. Token do chương trình tự lấy sau
    # khi CAPTCHA thành công và chỉ tồn tại trong bộ nhớ của phiên chạy.
    $username = Get-EnvValue $values 'GDT_USERNAME' ''
    $password = Get-EnvValue $values 'GDT_PASSWORD' ''
    $hasUsername = -not [string]::IsNullOrWhiteSpace($username)
    $hasPassword = -not [string]::IsNullOrWhiteSpace($password)
    if (-not $hasUsername -or -not $hasPassword) {
        throw 'Thiếu thông tin đăng nhập: cần điền cả GDT_USERNAME và GDT_PASSWORD.'
    }
    $token = ''

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

    # Tải XML song song có kiểm soát: XML_CONCURRENCY là số kết nối khởi đầu,
    # XML_MAX_CONCURRENCY là trần để cơ chế tự phục hồi tăng dần sau rate-limit,
    # XML_REQUEST_INTERVAL_MS là giãn cách giữa hai request XML liên tiếp.
    $xmlConcurrency = [int](Get-EnvValue $values 'XML_CONCURRENCY' '2')
    $xmlMaxConcurrency = [int](Get-EnvValue $values 'XML_MAX_CONCURRENCY' '3')
    $xmlRequestIntervalMs = [int](Get-EnvValue $values 'XML_REQUEST_INTERVAL_MS' '800')
    $browserUserAgent = (Get-EnvValue $values 'BROWSER_USER_AGENT' '').Trim()
    $logHttpProfile = ConvertTo-EnvBoolean 'LOG_HTTP_PROFILE' (Get-EnvValue $values 'LOG_HTTP_PROFILE' 'false')
    if ($xmlConcurrency -lt 1 -or $xmlConcurrency -gt 10) { throw 'XML_CONCURRENCY phải nằm trong khoảng 1-10.' }
    if ($xmlMaxConcurrency -lt 1 -or $xmlMaxConcurrency -gt 10) { throw 'XML_MAX_CONCURRENCY phải nằm trong khoảng 1-10.' }
    if ($xmlConcurrency -gt $xmlMaxConcurrency) { throw 'XML_CONCURRENCY không được lớn hơn XML_MAX_CONCURRENCY.' }
    if ($xmlRequestIntervalMs -lt 0 -or $xmlRequestIntervalMs -gt 60000) { throw 'XML_REQUEST_INTERVAL_MS phải nằm trong khoảng 0-60000.' }

    $includeRegular = ConvertTo-EnvBoolean 'INCLUDE_REGULAR' (Get-EnvValue $values 'INCLUDE_REGULAR' 'true')
    $includeSco = ConvertTo-EnvBoolean 'INCLUDE_SCO' (Get-EnvValue $values 'INCLUDE_SCO' 'true')
    if (-not $includeRegular -and -not $includeSco) { throw 'Phải bật ít nhất một trong INCLUDE_REGULAR hoặc INCLUDE_SCO.' }
    $fetchRelated = ConvertTo-EnvBoolean 'FETCH_RELATED' (Get-EnvValue $values 'FETCH_RELATED' 'true')

    $outputDirectory = Resolve-RepositoryPath $RepositoryRoot (Get-EnvValue $values 'OUTPUT_DIR' 'output')
    $outputName = Get-EnvValue $values 'OUTPUT_XLSX' 'HoaDonDienTu.xlsx'
    if ([System.IO.Path]::GetExtension($outputName).ToLowerInvariant() -ne '.xlsx') {
        throw 'OUTPUT_XLSX phải có phần mở rộng .xlsx.'
    }

    # Workbook .xlsx có sheet 'LinkTraCuu' để người dùng tự thêm/sửa link tra
    # cứu; file này được nạp khi xuất Excel.  Bỏ trống thì dùng bảng gốc.
    $lookupTableXlsx = ''
    $lookupTableValue = (Get-EnvValue $values 'LOOKUP_TABLE_XLSX' '').Trim()
    if (-not [string]::IsNullOrWhiteSpace($lookupTableValue)) {
        if ([System.IO.Path]::GetExtension($lookupTableValue).ToLowerInvariant() -ne '.xlsx') {
            throw 'LOOKUP_TABLE_XLSX phải là file .xlsx có sheet LinkTraCuu.'
        }
        $lookupTableXlsx = Resolve-RepositoryPath $RepositoryRoot $lookupTableValue
    }

    $proxyUrl = (Get-EnvValue $values 'PROXY_URL' '').Trim()
    $proxyUsername = Get-EnvValue $values 'PROXY_USERNAME' ''
    $proxyPassword = Get-EnvValue $values 'PROXY_PASSWORD' ''
    # Accept the convenient host:port:username:password form as well as a URL.
    if ($proxyUrl -notmatch '^[A-Za-z][A-Za-z0-9+.-]*://' -and $proxyUrl -match '^([^:]+):(\d+):([^:]+):(.*)$') {
        if ([string]::IsNullOrWhiteSpace($proxyUsername)) { $proxyUsername = $Matches[3] }
        if ([string]::IsNullOrWhiteSpace($proxyPassword)) { $proxyPassword = $Matches[4] }
        $proxyUrl = 'http://{0}:{1}' -f $Matches[1], $Matches[2]
    }
    elseif ($proxyUrl -notmatch '^[A-Za-z][A-Za-z0-9+.-]*://' -and $proxyUrl -match '^([^:]+):(\d+)$') {
        $proxyUrl = 'http://{0}:{1}' -f $Matches[1], $Matches[2]
    }
    $proxyUri = $null
    if (-not [string]::IsNullOrWhiteSpace($proxyUrl)) {
        $parsedProxyUri = $null
        if (-not [Uri]::TryCreate($proxyUrl, [UriKind]::Absolute, [ref]$parsedProxyUri)) {
            throw 'PROXY_URL không hợp lệ; ví dụ: http://127.0.0.1:8080.'
        }
        $proxyUri = $parsedProxyUri
        if ($proxyUri.Scheme -notin @('http', 'https') -or [string]::IsNullOrWhiteSpace($proxyUri.Host) -or $proxyUri.Port -lt 1) {
            throw 'PROXY_URL phải là URL http:// hoặc https:// có host và port.'
        }
        if (-not [string]::IsNullOrWhiteSpace($proxyUri.UserInfo) -or -not [string]::IsNullOrWhiteSpace($proxyUri.Query) -or -not [string]::IsNullOrWhiteSpace($proxyUri.Fragment)) {
            throw 'PROXY_URL không được chứa thông tin đăng nhập, query hoặc fragment; dùng PROXY_USERNAME/PROXY_PASSWORD.'
        }
    }
    if ([string]::IsNullOrWhiteSpace($proxyUsername) -and -not [string]::IsNullOrWhiteSpace($proxyPassword)) {
        throw 'PROXY_PASSWORD cần PROXY_USERNAME tương ứng.'
    }

    return [pscustomobject]@{
        Token = $token
        Username = $username.Trim()
        Password = $password
        BaseUrl = 'https://hoadondientu.gdt.gov.vn/api'
        ProxyUri = $proxyUri
        ProxyUsername = $proxyUsername
        ProxyPassword = $proxyPassword
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
        XmlConcurrency = $xmlConcurrency
        XmlMaxConcurrency = $xmlMaxConcurrency
        XmlRequestIntervalMs = $xmlRequestIntervalMs
        BrowserUserAgent = $browserUserAgent
        LogHttpProfile = $logHttpProfile
        IncludeRegular = $includeRegular
        IncludeSco = $includeSco
        FetchRelated = $fetchRelated
        RedownloadXml = ConvertTo-EnvBoolean 'REDOWNLOAD_XML' (Get-EnvValue $values 'REDOWNLOAD_XML' 'false')
        OverwriteOutput = ConvertTo-EnvBoolean 'OVERWRITE_OUTPUT' (Get-EnvValue $values 'OVERWRITE_OUTPUT' 'false')
        OutputDirectory = $outputDirectory
        OutputWorkbook = Join-Path $outputDirectory $outputName
        LookupTableXlsx = $lookupTableXlsx
        XmlDirectory = Join-Path $outputDirectory 'xml'
        LogDirectory = Join-Path $outputDirectory 'logs'
    }
}
