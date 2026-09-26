Set-StrictMode -Version 2.0

# Thời điểm bắt đầu request gần nhất trong toàn tiến trình. Áp dụng chung cho
# request danh sách, phân trang và tải XML để tránh dồn request vào máy chủ.
$script:LastGdtRequestUtc = [datetime]::MinValue
$script:AdaptiveRequestDelayMs = $null
$script:AdaptiveBaseDelayMs = $null
$script:SuccessfulRequestStreak = 0
$script:GdtWebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# Gate dùng chung cho tải XML đa luồng (runspace). $null = chạy tuần tự như cũ.
# Khi có gate, mọi runspace tham chiếu CÙNG một đối tượng in-process: giãn cách
# request, mốc tạm dừng 429, số luồng hiện hành và cờ dừng đều dùng chung.
$script:HddtSharedGate = $null

# Điều khiển dừng an toàn (tương ứng GdtStopRequested trong VBA).
# Ctrl+C lần 1 yêu cầu dừng sau request hiện tại, lần 2 dừng ngay.
# Dữ liệu đã tải trước khi dừng vẫn được xuất ra Excel.
$script:HddtStopRequested = $false
$script:LastGdtRequestAttempts = 0
$script:LastGdtStatusCode = 0
$script:LastGdtRequestUri = ''
$script:LastGdtRetryAfterSeconds = 0

function Reset-HddtStopRequest {
    $script:HddtStopRequested = $false
}

function Set-HddtStopRequest {
    $script:HddtStopRequested = $true
    if ($null -ne $script:HddtSharedGate) { $script:HddtSharedGate.StopRequested = $true }
}

function Test-HddtStopRequested {
    if ($null -ne $script:HddtSharedGate) { return [bool]$script:HddtSharedGate.StopRequested }
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
            if ($null -ne $script:HddtSharedGate) { $script:HddtSharedGate.StopRequested = $true }
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

# ---------------------------------------------------------------------------
# Shared gate cho tải XML đa luồng
# ---------------------------------------------------------------------------

function New-HddtSharedGate {
    param(
        [Parameter(Mandatory = $true)][int]$MaxWorkers,
        [Parameter(Mandatory = $true)][int]$BaseDelayMs
    )
    $max = [Math]::Max(1, $MaxWorkers)
    return [pscustomobject]@{
        Lock = New-Object System.Object
        LogLock = New-Object System.Object
        LoginLock = New-Object System.Object
        LastRequestUtc = [datetime]::MinValue
        PauseUntilUtc = [datetime]::MinValue
        MaxWorkers = $max
        CurrentLimit = $max
        ActiveWorkers = 0
        NextIndex = 0
        SuccessStreak = 0
        BaseDelayMs = [int]$BaseDelayMs
        AdaptiveDelayMs = [int]$BaseDelayMs
        StopRequested = $false
    }
}

function Set-HddtSharedGate {
    param($Gate)
    $script:HddtSharedGate = $Gate
}

function Lock-HddtGate {
    param([Parameter(Mandatory = $true)]$Gate)
    [System.Threading.Monitor]::Enter($Gate.Lock)
}

function Unlock-HddtGate {
    param([Parameter(Mandatory = $true)]$Gate)
    [System.Threading.Monitor]::Exit($Gate.Lock)
}

# Tạm dừng toàn cục mọi luồng (dùng sau HTTP 429) cho tới mốc thời gian mới.
# Chỉ đẩy mốc ra xa hơn, không bao giờ kéo gần lại.
function Set-GdtSharedPause {
    param([Parameter(Mandatory = $true)][int]$Seconds)
    $gate = $script:HddtSharedGate
    if ($null -eq $gate) { return }
    Lock-HddtGate $gate
    try {
        $until = [datetime]::UtcNow.AddSeconds($Seconds)
        if ($until -gt $gate.PauseUntilUtc) { $gate.PauseUntilUtc = $until }
    }
    finally { Unlock-HddtGate $gate }
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
    $gate = $script:HddtSharedGate
    if ($null -ne $gate) {
        # Quá tải: giảm số luồng (chia đôi, sàn 1) và tăng giãn cách. Các luồng
        # khác sẽ thấy CurrentLimit mới trước khi nhận hóa đơn kế tiếp.
        $limit = 1
        $delay = 0
        Lock-HddtGate $gate
        try {
            $gate.AdaptiveDelayMs = [Math]::Min(10000, [Math]::Max(1000, [int]$gate.AdaptiveDelayMs * 2))
            $gate.CurrentLimit = [Math]::Max(1, [int][Math]::Floor([int]$gate.CurrentLimit / 2))
            $gate.SuccessStreak = 0
            $limit = [int]$gate.CurrentLimit
            $delay = [int]$gate.AdaptiveDelayMs
        }
        finally { Unlock-HddtGate $gate }
        Write-HddtLog WARN ('[MẠNG] GDT giới hạn tốc độ; giảm còn {0} luồng, giãn cách {1} ms/request.' -f $limit, $delay)
        return
    }
    if (-not (Test-AdaptiveThrottleEnabled $Config)) { return }
    $currentDelay = Get-GdtRequestDelayMs $Config
    $script:AdaptiveRequestDelayMs = [Math]::Min(10000, [Math]::Max(1000, $currentDelay * 2))
    $script:SuccessfulRequestStreak = 0
    Write-HddtLog WARN ('[MẠNG] GDT giới hạn tốc độ; tự tăng giãn cách lên {0} ms/request.' -f $script:AdaptiveRequestDelayMs)
}

function Register-GdtRequestSuccess {
    param([Parameter(Mandatory = $true)]$Config)
    $gate = $script:HddtSharedGate
    if ($null -ne $gate) {
        # Ổn định trở lại: sau một chuỗi thành công, tăng số luồng lên 1 bậc
        # (tới trần), hoặc giảm giãn cách nếu đã ở trần.
        $raised = $false
        $decayed = $false
        $limit = 1
        $delay = 0
        Lock-HddtGate $gate
        try {
            $gate.SuccessStreak = [int]$gate.SuccessStreak + 1
            if ([int]$gate.SuccessStreak -ge 15) {
                $gate.SuccessStreak = 0
                if ([int]$gate.CurrentLimit -lt [int]$gate.MaxWorkers) {
                    $gate.CurrentLimit = [int]$gate.CurrentLimit + 1
                    $raised = $true
                }
                elseif ([int]$gate.AdaptiveDelayMs -gt [int]$gate.BaseDelayMs) {
                    $gate.AdaptiveDelayMs = [Math]::Max([int]$gate.BaseDelayMs, [int][Math]::Floor([int]$gate.AdaptiveDelayMs * 0.85))
                    $decayed = $true
                }
            }
            $limit = [int]$gate.CurrentLimit
            $delay = [int]$gate.AdaptiveDelayMs
        }
        finally { Unlock-HddtGate $gate }
        if ($raised) { Write-HddtLog INFO ('[MẠNG] Kết nối ổn định; tăng lên {0} luồng tải XML.' -f $limit) }
        elseif ($decayed) { Write-HddtLog INFO ('[MẠNG] Kết nối ổn định; giảm giãn cách xuống {0} ms/request.' -f $delay) }
        return
    }
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
    $gate = $script:HddtSharedGate
    if ($null -ne $gate) {
        # Đặt trước "khe" khởi request kế tiếp trong khi giữ khóa: các luồng
        # nhận mốc thời gian tăng dần, nên khoảng cách tối thiểu giữa các
        # request vẫn được giữ dù chạy song song.
        $waitMs = 0
        Lock-HddtGate $gate
        try {
            $now = [datetime]::UtcNow
            $target = $now
            if ($gate.LastRequestUtc -ne [datetime]::MinValue) {
                $spacing = [datetime]$gate.LastRequestUtc + [timespan]::FromMilliseconds([int]$gate.AdaptiveDelayMs)
                if ($spacing -gt $target) { $target = $spacing }
            }
            if ($gate.PauseUntilUtc -gt $target) { $target = $gate.PauseUntilUtc }
            $gate.LastRequestUtc = $target
            $waitMs = [int][Math]::Ceiling(($target - $now).TotalMilliseconds)
        }
        finally { Unlock-HddtGate $gate }
        if ($waitMs -gt 0) {
            Write-HddtLog DEBUG ('Giãn cách request (đa luồng): chờ {0} ms.' -f $waitMs)
            Start-Sleep -Milliseconds $waitMs
        }
        return
    }
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
        [switch]$SkipAuthorization
    )

    if (-not $Uri.StartsWith($Config.BaseUrl + '/', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Từ chối gửi token tới URL ngoài máy chủ GDT đã cố định.'
    }

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $attempt = 0
    $rateLimitAttempts = 0
    $requestToken = ''
    Reset-GdtRequestContext
    $script:LastGdtRequestUri = [string]$Uri
    while ($true) {
        if (Test-HddtStopRequested) { throw 'Đã dừng theo yêu cầu; không gửi request mới.' }
        $script:LastGdtRetryAfterSeconds = 0
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
                $requestToken = [string]$tokenProperty.Value
                $headers['Authorization'] = 'Bearer ' + $requestToken
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
            $tokenProperty = $Config.PSObject.Properties['Token']
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
                # Đăng nhập bằng tài khoản: tự đăng nhập lại lấy token mới rồi
                # thử lại request này, thay vì bỏ phiên tải giữa chừng.
                $attempt++
                $script:LastGdtRequestAttempts = $attempt
                $gate = $script:HddtSharedGate
                if ($null -ne $gate) {
                    # Đa luồng: chỉ một luồng đăng nhập lại (giữ LoginLock);
                    # các luồng khác chờ khóa rồi dùng token mới, không cùng lúc
                    # giải CAPTCHA hay vượt mặt nhau.
                    [System.Threading.Monitor]::Enter($gate.LoginLock)
                    try {
                        if ([string]$Config.Token -eq $requestToken) {
                            Write-HddtLog WARN ('[MẠNG] HTTP {0}; tự đăng nhập lại rồi thử tiếp.' -f $status)
                            try {
                                $newToken = Request-GdtSessionToken -Config $Config
                                $Config.Token = $newToken
                            }
                            catch {
                                throw "Tự đăng nhập lại sau HTTP $status thất bại: $($_.Exception.Message)"
                            }
                        }
                    }
                    finally { [System.Threading.Monitor]::Exit($gate.LoginLock) }
                    continue
                }
                Write-HddtLog WARN ('[MẠNG] HTTP {0}; tự đăng nhập lại rồi thử tiếp.' -f $status)
                try {
                    $newToken = Request-GdtSessionToken -Config $Config
                    $Config.Token = $newToken
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
                Register-GdtRateLimit $Config
                # 429 được thử lại độc lập với MAX_RETRIES: đợt rate-limit của
                # GDT có thể kéo dài nhiều phút, phiên tải không nên dừng chỉ
                # vì hết 4 lần thử ngắn. Trần tối đa là Get-429RetryLimit.
                $rateLimitAttempts++
                if ($rateLimitAttempts -gt (Get-429RetryLimit)) {
                    throw ('GDT vẫn giới hạn tốc độ sau {0} lần thử lại. Hãy chờ vài phút hoặc tăng REQUEST_DELAY_MS rồi chạy lại; dữ liệu đã tải vẫn được xuất ra Excel.' -f (Get-429RetryLimit))
                }
                $retryAfterSeconds = Get-HttpRetryAfterSeconds $_
                $script:LastGdtRetryAfterSeconds = $retryAfterSeconds
                $waitSeconds = Get-RetryDelaySeconds -StatusCode 429 -Attempt $rateLimitAttempts -RetryAfterSeconds $retryAfterSeconds
                Set-GdtSharedPause -Seconds $waitSeconds
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
