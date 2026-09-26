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
# Trong phase tải XML song song, cờ dừng sống trong trạng thái dùng chung để
# mọi worker runspace nhìn thấy cùng một giá trị.
$script:HddtStopRequested = $false
$script:LastGdtRequestAttempts = 0
$script:LastGdtStatusCode = 0
$script:LastGdtRequestUri = ''
$script:LastGdtRetryAfterSeconds = 0

# Chỉ đúng khi src/XmlScheduler.ps1 được nạp; khi đó request ExportXml đi qua
# bộ điều tiết XML (Enter/Exit-GdtXmlRequestSlot) thay vì giãn cách tổng.
$script:HddtXmlThrottleAvailable = $false

function Test-GdtXmlThrottleAvailable {
    return [bool]$script:HddtXmlThrottleAvailable
}

function Reset-HddtStopRequest {
    $script:HddtStopRequested = $false
    Set-HddtSharedValue -Key 'StopRequested' -Value $false
}

function Set-HddtStopRequest {
    $script:HddtStopRequested = $true
    Set-HddtSharedValue -Key 'StopRequested' -Value $true
}

function Test-HddtStopRequested {
    if ([bool]$script:HddtStopRequested) { return $true }
    return [bool](Get-HddtSharedValue -Key 'StopRequested' -Default $false)
}

# Số lần thử và HTTP status của request GDT gần nhất, dùng để tạo thông báo
# lỗi cho chuỗi hóa đơn liên quan (tương ứng LastGdtAttempts/LastGdtStatus trong VBA).
function Get-GdtLastRequestAttempts {
    return [int]$script:LastGdtRequestAttempts
}

function Get-GdtLastStatusCode {
    return [int]$script:LastGdtStatusCode
}

function Get-GdtLastRequestUri {
    return [string]$script:LastGdtRequestUri
}

function Get-GdtLastRetryAfterSeconds {
    return [int]$script:LastGdtRetryAfterSeconds
}

function Reset-GdtRequestContext {
    $script:LastGdtRequestAttempts = 0
    $script:LastGdtStatusCode = 0
    $script:LastGdtRequestUri = ''
    $script:LastGdtRetryAfterSeconds = 0
}

# Token hiện hành: ưu tiên trạng thái dùng chung khi đang chạy pipeline tải
# song song, ngược lại lấy từ object cấu hình (phase danh sách chạy tuần tự).
function Get-GdtAuthToken {
    param($Config)
    $token = [string](Get-HddtSharedValue -Key 'Token' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($token)) { return $token }
    if ($null -eq $Config) { return '' }
    $property = $Config.PSObject.Properties['Token']
    if ($null -eq $property -or $null -eq $property.Value) { return '' }
    return [string]$property.Value
}

function Set-GdtAuthToken {
    param($Config, [Parameter(Mandatory = $true)][string]$Token)
    Set-HddtSharedValue -Key 'Token' -Value $Token
    if ($null -ne $Config) {
        $property = $Config.PSObject.Properties['Token']
        if ($null -ne $property) { $property.Value = $Token }
    }
}

# Làm mới token theo kiểu single-flight: chỉ một luồng đăng nhập lại, các
# luồng khác chờ token đổi thay vì mỗi em giành nhau gọi CAPTCHA/login song
# song. Không có pipeline (chạy tuần tự) thì giữ nguyên hành vi cũ.
function Invoke-GdtTokenRefresh {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [string]$FailedToken = '',
        [int]$TimeoutSeconds = 180
    )

    $shared = Get-HddtSharedState
    if ($null -eq $shared) {
        return Request-GdtSessionToken -Config $Config
    }

    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    if ([Threading.Monitor]::TryEnter($shared.SyncRoot, 0)) {
        try {
            $current = [string](Get-HddtSharedValue -Key 'Token' -Default '')
            if (-not [string]::IsNullOrWhiteSpace($current) -and $current -ne $FailedToken) {
                # Luồng khác đã làm mới xong trong lúc mình chờ vào lock.
                return $current
            }
            Write-HddtLog DEBUG '[MẠNG] Đăng nhập lại để lấy token phiên mới.'
            $newToken = Request-GdtSessionToken -Config $Config
            $refreshCount = [int](Get-HddtSharedValue -Key 'AuthRefreshCount' -Default 0)
            Set-HddtSharedValue -Key 'AuthRefreshCount' -Value ($refreshCount + 1)
            Set-HddtSharedValue -Key 'Token' -Value $newToken
            return $newToken
        }
        finally { [Threading.Monitor]::Exit($shared.SyncRoot) }
    }

    # Đang có luồng khác làm mới: chờ token đổi, không gọi CAPTCHA lần nữa.
    while ([datetime]::UtcNow -lt $deadline) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không chờ làm mới token.' }
        $current = [string](Get-HddtSharedValue -Key 'Token' -Default '')
        if (-not [string]::IsNullOrWhiteSpace($current) -and $current -ne $FailedToken) { return $current }
        Start-Sleep -Milliseconds 250
    }
    throw 'Hết thời gian chờ làm mới token phiên (180s).'
}

function Register-HddtStopHandler {
    try {
        $null = [Console]::add_CancelKeyPress({
            param($sender, $eventArgs)
            if (Test-HddtStopRequested) {
                # Nhấn lần 2: cho phép PowerShell dừng ngay (không xuất Excel).
                $eventArgs.Cancel = $false
                return
            }
            Set-HddtStopRequest
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
        # 429 cần chờ dài: đợt rate-limit của GDT thường kéo dài vài phút, nghĩ
        # ngắn quá là quay lại dập request ngay khi vừa hết giãn cách.
        $calculated = [Math]::Min(600, 15 * [Math]::Pow(2, $Attempt - 1))
    }
    else {
        $calculated = [Math]::Min(30, [Math]::Pow(2, $Attempt))
    }
    return [int][Math]::Max($calculated, $RetryAfterSeconds)
}

function Get-429RetryLimit {
    # Số lần thử lại tối đa cho HTTP 429, độc lập với MAX_RETRIES. Đặt cao để
    # một đợt rate-limit dài không làm dừng phiên tải: đã gặp phiên dừng sau
    # 4 lần thử (tổng ~1,5 phút) dù GDT vẫn trả 429 liên tục nhiều phút.
    # Đặt MAX_RETRIES cao hơn nếu muốn, giới hạn này chỉ là trần tối đa.
    return 12
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

function Add-HddtProxyParameters {
    param([hashtable]$Parameters, [Parameter(Mandatory = $true)]$Config)
    $proxyProperty = $Config.PSObject.Properties['ProxyUri']
    if ($null -eq $proxyProperty -or $null -eq $proxyProperty.Value) { return }
    $proxyUri = [Uri]$proxyProperty.Value
    $Parameters['Proxy'] = $proxyUri

    $usernameProperty = $Config.PSObject.Properties['ProxyUsername']
    $passwordProperty = $Config.PSObject.Properties['ProxyPassword']
    $username = if ($null -ne $usernameProperty) { [string]$usernameProperty.Value } else { '' }
    $password = if ($null -ne $passwordProperty) { [string]$passwordProperty.Value } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($username)) {
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        $Parameters['ProxyCredential'] = New-Object System.Management.Automation.PSCredential($username, $securePassword)
    }
    Write-HddtLog DEBUG ('[MẠNG] Dùng proxy {0}:{1}.' -f $proxyUri.Host, $proxyUri.Port)
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
        [switch]$SkipAuthorization,
        # Header profile cho endpoint (Captcha/Login/InvoiceQuery/
        # InvoiceRelation/ExportXml); để trống thì suy ra từ URI.
        [string]$RequestProfile = ''
    )

    if (-not $Uri.StartsWith($Config.BaseUrl + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Từ chối gửi token tới URL ngoài máy chủ GDT đã cố định.'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $profileName = Resolve-GdtRequestProfile -Uri $Uri -RequestProfile $RequestProfile
    $isXmlDownload = ($profileName -eq 'ExportXml') -and (Test-GdtXmlThrottleAvailable)
    $attempt = 0
    $rateLimitAttempts = 0
    Reset-GdtRequestContext
    $script:LastGdtRequestUri = [string]$Uri
    while ($true) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không gửi request mới.' }
        $script:LastGdtRetryAfterSeconds = 0
        $slotHeld = $false
        try {
            if ($isXmlDownload) {
                # Request XML đi qua bộ điều tiết riêng: chặn theo cooldown
                # rate-limit, số kết nối đang mở và giãn cách giữa hai request.
                $null = Enter-GdtXmlRequestSlot -Config $Config
                $slotHeld = $true
            }
            else {
                Wait-GdtRequestSlot -Config $Config
            }
            $script:LastGdtRequestUtc = [datetime]::UtcNow
            $script:LastGdtRequestAttempts = $attempt + 1
            $requestWatch = [Diagnostics.Stopwatch]::StartNew()
            $authorizationToken = ''
            if (-not $SkipAuthorization) { $authorizationToken = Get-GdtAuthToken -Config $Config }
            $headers = Get-GdtRequestHeaders -RequestProfile $profileName -Config $Config `
                -AuthorizationToken $authorizationToken -AdditionalHeaders $ExtraHeaders
            $headers['Request-Id'] = [guid]::NewGuid().ToString()
            if ($AsBytes) {
                $headers['Accept'] = 'application/zip, application/xml, application/octet-stream, */*'
            }
            # Token đã dùng cho request này: nếu bị 401 thì đó là token cũ để
            # phân biệt với token do luồng khác vừa làm mới.
            $tokenUsed = $authorizationToken
            $parameters = @{
                Uri = $Uri
                Method = $Method
                Headers = $headers
                TimeoutSec = $Config.HttpTimeoutSeconds
                UseBasicParsing = $true
                WebSession = $script:GdtWebSession
                ErrorAction = 'Stop'
            }
            Add-HddtProxyParameters -Parameters $parameters -Config $Config
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
            if ($isXmlDownload) {
                Exit-GdtXmlRequestSlot -ResponseTimeMs ([int]$requestWatch.ElapsedMilliseconds)
                $slotHeld = $false
                Register-GdtXmlSuccess
            }
            else {
                Register-GdtRequestSuccess $Config
            }
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
            if ($slotHeld) {
                Exit-GdtXmlRequestSlot
                $slotHeld = $false
            }
            $status = Get-HttpStatusCode $_
            $script:LastGdtStatusCode = $status
            $usernameProperty = $Config.PSObject.Properties['Username']
            $passwordProperty = $Config.PSObject.Properties['Password']
            $hasCredentials = ($null -ne $usernameProperty -and -not [string]::IsNullOrWhiteSpace([string]$usernameProperty.Value) -and $null -ne $passwordProperty -and -not [string]::IsNullOrWhiteSpace([string]$passwordProperty.Value))
            if ($status -eq 401 -or $status -eq 403) {
                if ($SkipAuthorization) {
                    if ($Uri -match '/security-taxpayer/authenticate') {
                        throw "Đăng nhập GDT không thành công (HTTP $status). Kiểm tra GDT_USERNAME/GDT_PASSWORD hoặc quyền truy cập tài khoản."
                    }
                    throw "Yêu cầu CAPTCHA không được phép (HTTP $status)."
                }
                if (-not $hasCredentials) {
                    throw "Phiên đăng nhập không hợp lệ (HTTP $status)."
                }
                # Đăng nhập bằng tài khoản: làm mới token rồi thử lại request
                # này, thay vì bỏ phiên tải giữa chừng. Có trần MAX_RETRIES để
                # không lặp vô hạn khi token mới vẫn bị từ chối, và các luồng
                # tải song song dùng chung một lần đăng nhập (single-flight).
                $attempt++
                $script:LastGdtRequestAttempts = $attempt
                if ($attempt -gt $Config.MaxRetries) {
                    throw ('Phiên đăng nhập không hợp lệ (HTTP {0}) sau {1} lần làm mới token.' -f $status, $Config.MaxRetries)
                }
                Write-HddtLog WARN ('[MẠNG] HTTP {0}; tự đăng nhập lại rồi thử tiếp.' -f $status)
                try {
                    $newToken = Invoke-GdtTokenRefresh -Config $Config -FailedToken $tokenUsed
                    Set-GdtAuthToken -Config $Config -Token $newToken
                }
                catch {
                    throw "Tự đăng nhập lại sau HTTP $status thất bại: $($_.Exception.Message)"
                }
                continue
            }
            if ($AsBytes -and $status -eq 500) {
                # Endpoint export-xml dùng HTTP 500 khi hóa đơn không có hồ sơ
                # XML gốc. VBA nguồn cũng bỏ qua ngay trường hợp này.
                throw 'GDT không có hồ sơ XML gốc cho hóa đơn này (HTTP 500).'
            }
            if ($status -eq 429) {
                # 429 được thử lại độc lập với MAX_RETRIES: đợt rate-limit của
                # GDT có thể kéo dài nhiều phút, phiên tải không nên dừng chỉ
                # vì hết 4 lần thử ngắn. Trần tối đa là Get-429RetryLimit.
                $retryAfterSeconds = Get-HttpRetryAfterSeconds $_
                $script:LastGdtRetryAfterSeconds = $retryAfterSeconds
                $rateLimitAttempts++
                if ($rateLimitAttempts -gt (Get-429RetryLimit)) {
                    throw ('GDT vẫn giới hạn tốc độ sau {0} lần thử lại. Hãy chờ vài phút hoặc tăng REQUEST_DELAY_MS rồi chạy lại; dữ liệu đã tải vẫn được xuất ra Excel.' -f (Get-429RetryLimit))
                }
                if ($isXmlDownload) {
                    # Request XML: không ngủ local. Enter-GdtXmlRequestSlot sẽ
                    # chặn đúng theo cooldown, giảm số kết nối và tăng giãn cách.
                    Register-GdtXmlRateLimit -RetryAfterSeconds $retryAfterSeconds
                    Write-HddtLog WARN ("[MẠNG] HTTP 429; thử lại {0}/{1} sau cooldown tải XML (độc lập MAX_RETRIES)." -f $rateLimitAttempts, (Get-429RetryLimit))
                    if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; ngưng thử lại.' }
                    continue
                }
                Register-GdtRateLimit $Config
                $waitSeconds = Get-RetryDelaySeconds -StatusCode 429 -Attempt $rateLimitAttempts -RetryAfterSeconds $retryAfterSeconds
                Write-HddtLog WARN ("[MẠNG] HTTP 429; thử lại {0}/{1} sau {2}s (độc lập MAX_RETRIES)." -f $rateLimitAttempts, (Get-429RetryLimit), $waitSeconds)
                if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; ngưng thử lại.' }
                Start-Sleep -Seconds $waitSeconds
                continue
            }

            $retryable = ($status -eq 0 -or $status -ge 500)
            if (-not $retryable -or $attempt -ge $Config.MaxRetries) {
                if ($status -gt 0) { throw "Yêu cầu GDT thất bại (HTTP $status)." }
                throw "Không kết nối được tới GDT: $($_.Exception.Message)"
            }

            $attempt++
            $waitSeconds = Get-RetryDelaySeconds -StatusCode $status -Attempt $attempt
            Write-HddtLog WARN ("[MẠNG] HTTP {0}; thử lại {1}/{2} sau {3}s." -f $status, $attempt, $Config.MaxRetries, $waitSeconds)
            if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; ngưng thử lại.' }
            Start-Sleep -Seconds $waitSeconds
        }
    }
}
