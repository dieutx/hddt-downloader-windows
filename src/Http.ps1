Set-StrictMode -Version 2.0

# Thời điểm bắt đầu request gần nhất trong toàn tiến trình. Áp dụng chung cho
# request danh sách, phân trang và tải XML để tránh dồn request vào máy chủ.
$script:LastGdtRequestUtc = [datetime]::MinValue
$script:AdaptiveRequestDelayMs = $null
$script:AdaptiveBaseDelayMs = $null
$script:SuccessfulRequestStreak = 0
$script:GdtWebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# Điều khiển dừng an toàn (tương ứng GdtStopRequested trong VBA).
# Ctrl+C lần 1 yêu cầu dừng sau request hiện tại, lần 2 dừng ngay.
# Dữ liệu đã tải trước khi dừng vẫn được xuất ra Excel.
$script:HddtStopRequested = $false
$script:LastGdtRequestAttempts = 0
$script:LastGdtStatusCode = 0

function Reset-HddtStopRequest {
    $script:HddtStopRequested = $false
}

function Set-HddtStopRequest {
    $script:HddtStopRequested = $true
}

function Test-HddtStopRequested {
    return [bool]$script:HddtStopRequested
}

# Số lần thử và HTTP status của request GDT gần nhất, dùng để tạo thông báo
# lỗi cho chuỗi hóa đơn liên quan (tương ứng LastGdtAttempts/LastGdtStatus trong VBA).
function Get-GdtLastRequestAttempts {
    return [int]$script:LastGdtRequestAttempts
}

function Get-GdtLastStatusCode {
    return [int]$script:LastGdtStatusCode
}

function Register-HddtStopHandler {
    try {
        $null = [Console]::add_CancelKeyPress({
            param($sender, $eventArgs)
            if ($script:HddtStopRequested) {
                # Nhấn lần 2: cho phép PowerShell dừng ngay (không xuất Excel).
                $eventArgs.Cancel = $false
                return
            }
            $script:HddtStopRequested = $true
            $eventArgs.Cancel = $true
            try {
                Write-Host ''
                Write-Host '[STOP] Đã yêu cầu dừng (Ctrl+C); chờ request hiện tại hoàn tất. Nhấn Ctrl+C lần nữa để dừng ngay.' -ForegroundColor Yellow
            } catch { }
        })
        return $true
    }
    catch { return $false }
}

function Test-AdaptiveThrottleEnabled {
    param([Parameter(Mandatory = $true)]$Config)
    $property = $Config.PSObject.Properties['AdaptiveThrottle']
    return ($null -ne $property -and [bool]$property.Value)
}

function Get-GdtRequestDelayMs {
    param([Parameter(Mandatory = $true)]$Config)
    $baseDelay = [int]$Config.RequestDelayMs
    if (-not (Test-AdaptiveThrottleEnabled $Config)) { return $baseDelay }
    if ($null -eq $script:AdaptiveRequestDelayMs -or $script:AdaptiveBaseDelayMs -ne $baseDelay) {
        $script:AdaptiveBaseDelayMs = $baseDelay
        $script:AdaptiveRequestDelayMs = $baseDelay
        $script:SuccessfulRequestStreak = 0
    }
    return [int]$script:AdaptiveRequestDelayMs
}

function Register-GdtRateLimit {
    param([Parameter(Mandatory = $true)]$Config)
    if (-not (Test-AdaptiveThrottleEnabled $Config)) { return }
    $currentDelay = Get-GdtRequestDelayMs $Config
    $script:AdaptiveRequestDelayMs = [Math]::Min(10000, [Math]::Max(1000, $currentDelay * 2))
    $script:SuccessfulRequestStreak = 0
    Write-HddtLog WARN ('[MẠNG] GDT giới hạn tốc độ; tự tăng giãn cách lên {0} ms/request.' -f $script:AdaptiveRequestDelayMs)
}

function Register-GdtRequestSuccess {
    param([Parameter(Mandatory = $true)]$Config)
    if (-not (Test-AdaptiveThrottleEnabled $Config)) { return }
    $currentDelay = Get-GdtRequestDelayMs $Config
    if ($currentDelay -le $Config.RequestDelayMs) { return }
    $script:SuccessfulRequestStreak++
    if ($script:SuccessfulRequestStreak -ge 20) {
        $script:AdaptiveRequestDelayMs = [Math]::Max([int]$Config.RequestDelayMs, [int][Math]::Floor($currentDelay * 0.85))
        $script:SuccessfulRequestStreak = 0
        Write-HddtLog INFO ('[MẠNG] Kết nối ổn định; giảm giãn cách xuống {0} ms/request.' -f $script:AdaptiveRequestDelayMs)
    }
}

function Get-HttpStatusCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        $response = $ErrorRecord.Exception.Response
        if ($null -ne $response) {
            $statusProperty = $response.PSObject.Properties['StatusCode']
            if ($null -ne $statusProperty -and $null -ne $statusProperty.Value) {
                return [int]$statusProperty.Value
            }
        }
    }
    catch { }
    return 0
}

function Get-HttpRetryAfterSeconds {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        if ($null -eq $ErrorRecord.Exception.Response) { return 0 }
        $value = [string]$ErrorRecord.Exception.Response.Headers['Retry-After']
        if ([string]::IsNullOrWhiteSpace($value)) { return 0 }

        $seconds = 0
        if ([int]::TryParse($value, [ref]$seconds)) { return [Math]::Max(0, $seconds) }

        $retryAt = [datetime]::MinValue
        if ([datetime]::TryParse($value, [ref]$retryAt)) {
            return [Math]::Max(0, [Math]::Ceiling(($retryAt.ToUniversalTime() - [datetime]::UtcNow).TotalSeconds))
        }
    }
    catch { }
    return 0
}

function Get-RetryDelaySeconds {
    param([int]$StatusCode, [int]$Attempt, [int]$RetryAfterSeconds = 0)
    if ($StatusCode -eq 429) {
        $calculated = [Math]::Min(120, 5 * [Math]::Pow(2, $Attempt - 1))
    }
    else {
        $calculated = [Math]::Min(30, [Math]::Pow(2, $Attempt))
    }
    return [int][Math]::Max($calculated, $RetryAfterSeconds)
}

function Wait-GdtRequestSlot {
    param([Parameter(Mandatory = $true)]$Config)
    $requestDelayMs = Get-GdtRequestDelayMs $Config
    if ($requestDelayMs -le 0 -or $script:LastGdtRequestUtc -eq [datetime]::MinValue) { return }

    $elapsedMs = ([datetime]::UtcNow - $script:LastGdtRequestUtc).TotalMilliseconds
    $remainingMs = [int][Math]::Ceiling($requestDelayMs - $elapsedMs)
    if ($remainingMs -gt 0) {
        Write-HddtLog DEBUG ('Giãn cách request: chờ {0} ms.' -f $remainingMs)
        Start-Sleep -Milliseconds $remainingMs
    }
}

function Invoke-GdtRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Uri,
        [switch]$AsBytes,
        [ValidateSet('Get', 'Post')][string]$Method = 'Get',
        [string]$Body,
        [string]$ContentType = 'application/json',
        [hashtable]$ExtraHeaders,
        [switch]$SkipAuthorization
    )

    if (-not $Uri.StartsWith($Config.BaseUrl + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Từ chối gửi token tới URL ngoài máy chủ GDT đã cố định.'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $attempt = 0
    $script:LastGdtRequestAttempts = 0
    $script:LastGdtStatusCode = 0
    while ($true) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không gửi request mới.' }
        try {
            Wait-GdtRequestSlot -Config $Config
            $script:LastGdtRequestUtc = [datetime]::UtcNow
            $script:LastGdtRequestAttempts = $attempt + 1
            $requestWatch = [Diagnostics.Stopwatch]::StartNew()
            $headers = @{
                Accept = if ($AsBytes) { 'application/zip, application/xml, */*' } else { 'application/json' }
                'Request-Id' = [guid]::NewGuid().ToString()
                'User-Agent' = 'HDDT-Windows-PowerShell/1.0'
            }
            $tokenProperty = $Config.PSObject.Properties['Token']
            if (-not $SkipAuthorization -and $null -ne $tokenProperty -and -not [string]::IsNullOrWhiteSpace([string]$tokenProperty.Value)) {
                $headers['Authorization'] = 'Bearer ' + [string]$tokenProperty.Value
            }
            if ($null -ne $ExtraHeaders) {
                foreach ($headerName in $ExtraHeaders.Keys) { $headers[$headerName] = $ExtraHeaders[$headerName] }
            }
            $parameters = @{
                Uri = $Uri
                Method = $Method
                Headers = $headers
                TimeoutSec = $Config.HttpTimeoutSeconds
                UseBasicParsing = $true
                WebSession = $script:GdtWebSession
                ErrorAction = 'Stop'
            }
            # Lưu ý: tham số [string] không truyền vào nhận giá trị '' chứ không phải
            # $null; kiểm tra rỗng để GET/DELETE không bị gửi kèm body (PowerShell
            # báo "Cannot send a content-body with this verb-type").
            if (-not [string]::IsNullOrEmpty($Body)) {
                $parameters['Body'] = $Body
                $parameters['ContentType'] = $ContentType
            }
            Write-HddtLog DEBUG ('HTTP {0} {1}' -f $Method.ToUpperInvariant(), $Uri)
            $response = Invoke-WebRequest @parameters
            $requestWatch.Stop()
            Register-GdtRequestSuccess $Config
            if ($AsBytes) {
                $binaryContent = $null
                $streamProperty = $response.PSObject.Properties['RawContentStream']
                if ($null -ne $streamProperty -and $null -ne $streamProperty.Value) {
                    $sourceStream = $streamProperty.Value
                    if ($sourceStream.CanSeek) { $sourceStream.Position = 0 }
                    $memoryStream = New-Object IO.MemoryStream
                    try {
                        $sourceStream.CopyTo($memoryStream)
                        $binaryContent = $memoryStream.ToArray()
                    }
                    finally { $memoryStream.Dispose() }
                }
                elseif ($response.Content -is [byte[]]) {
                    $binaryContent = [byte[]]$response.Content
                }
                if ($null -eq $binaryContent -or $binaryContent.Length -eq 0) {
                    throw 'Máy chủ trả về nội dung tải xuống rỗng hoặc không hợp lệ.'
                }
                Write-HddtLog DEBUG ('HTTP tải dữ liệu vào bộ nhớ thành công sau {0} ms; {1} byte.' -f $requestWatch.ElapsedMilliseconds, $binaryContent.Length)
                return ,$binaryContent
            }
            Write-HddtLog DEBUG ('HTTP {0} sau {1} ms.' -f [int]$response.StatusCode, $requestWatch.ElapsedMilliseconds)
            return [string]$response.Content
        }
        catch {
            $status = Get-HttpStatusCode $_
            $script:LastGdtStatusCode = $status
            if ($status -eq 401 -or $status -eq 403) {
                throw "Token hết hạn, không hợp lệ hoặc không có quyền (HTTP $status)."
            }
            if ($AsBytes -and $status -eq 500) {
                # Endpoint export-xml dùng HTTP 500 khi hóa đơn không có hồ sơ
                # XML gốc. VBA nguồn cũng bỏ qua ngay trường hợp này.
                throw 'GDT không có hồ sơ XML gốc cho hóa đơn này (HTTP 500).'
            }
            if ($status -eq 429) { Register-GdtRateLimit $Config }

            $retryable = ($status -eq 0 -or $status -eq 429 -or $status -ge 500)
            if (-not $retryable -or $attempt -ge $Config.MaxRetries) {
                if ($status -eq 429) {
                    throw 'GDT vẫn giới hạn tốc độ (HTTP 429). Hãy chờ vài phút hoặc tăng REQUEST_DELAY_MS rồi chạy lại.'
                }
                if ($status -gt 0) { throw "Yêu cầu GDT thất bại (HTTP $status)." }
                throw "Không kết nối được tới GDT: $($_.Exception.Message)"
            }

            $attempt++
            $retryAfterSeconds = Get-HttpRetryAfterSeconds $_
            $waitSeconds = Get-RetryDelaySeconds -StatusCode $status -Attempt $attempt -RetryAfterSeconds $retryAfterSeconds
            Write-HddtLog WARN ("[MẠNG] HTTP {0}; thử lại {1}/{2} sau {3}s." -f $status, $attempt, $Config.MaxRetries, $waitSeconds)
            if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; ngưng thử lại.' }
            Start-Sleep -Seconds $waitSeconds
        }
    }
}
